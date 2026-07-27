"""获取钉钉通讯录全部用户信息，并可选写入 ``sys_user`` / ``sys_user_oauth``。

钉钉不提供一次性拉全员接口，需：
1. 从根部门（dept_id=1）递归 ``topapi/v2/department/listsub`` 拿到全部部门；
2. 对每个部门分页调用 ``topapi/v2/user/list``；
3. 按 ``userid`` 去重（一人可属多部门）；
4. ``--save`` 时按 ``(oauth_type=dingtalk, open_id=userid)`` upsert 到 ``sys_user_oauth``
   （无绑定用户则先建/匹配 ``sys_user``）。

权限（企业内部应用）：
- 通讯录部门信息读权限（qyapi_get_department_list）
- 通讯录部门成员读权限（qyapi_get_department_member）

文档：
- https://open.dingtalk.com/document/orgapp/obtain-the-department-list-v2
- https://open.dingtalk.com/document/orgapp/queries-the-complete-information-of-a-department-user

运行（项目根目录）::

    python -m api.ding_disk.getUsers
    python -m api.ding_disk.getUsers --summary
    python -m api.ding_disk.getUsers --save
    python -m api.ding_disk.getUsers --save --dry-run
    python -m api.ding_disk.getUsers --out users.json
    python -m api.ding_disk.getUsers --dept-id 1 --no-recurse
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set

JsonDict = Dict[str, Any]
MappingLike = Dict[str, Any]

# 家校通讯录根，不属于企业内部通讯录
_SKIP_DEPT_IDS: Set[int] = {-7}

OAUTH_TYPE_DINGTALK = "dingtalk"
TABLE_USER = "sys_user"
TABLE_OAUTH = "sys_user_oauth"


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def list_all_departments(
    client: Any,
    *,
    root_dept_id: int = 1,
    language: str = "zh_CN",
    include_root: bool = True,
) -> List[JsonDict]:
    """递归获取授权范围内全部部门（BFS）。

    返回列表每项含 ``dept_id`` / ``name`` / ``parent_id`` 等；
    ``include_root=True`` 时额外插入根部门占位（name=根部门）。
    """
    departments: List[JsonDict] = []
    seen: Set[int] = set()
    queue: List[int] = [int(root_dept_id)]

    if include_root:
        departments.append(
            {
                "dept_id": int(root_dept_id),
                "name": "根部门",
                "parent_id": 0,
            }
        )
        seen.add(int(root_dept_id))

    while queue:
        parent_id = queue.pop(0)
        for dept in client.list_sub_departments(parent_id, language=language):
            try:
                dept_id = int(dept.get("dept_id"))
            except (TypeError, ValueError):
                continue
            if dept_id in _SKIP_DEPT_IDS or dept_id in seen:
                continue
            seen.add(dept_id)
            departments.append(dept)
            queue.append(dept_id)
    return departments


def iter_department_users(
    client: Any,
    dept_id: int,
    *,
    page_size: int = 100,
    language: str = "zh_CN",
    contain_access_limit: bool = False,
) -> Iterable[JsonDict]:
    """分页迭代单个部门下的用户详情（不含子部门）。"""
    cursor = 0
    while True:
        result = client.list_department_users(
            dept_id,
            cursor=cursor,
            size=page_size,
            language=language,
            contain_access_limit=contain_access_limit,
        )
        rows = result.get("list") or []
        if isinstance(rows, list):
            for row in rows:
                if isinstance(row, dict):
                    yield row
        if not result.get("has_more"):
            break
        next_cursor = result.get("next_cursor")
        if next_cursor is None:
            break
        cursor = int(next_cursor)


def get_all_users(
    client: Optional[Any] = None,
    *,
    root_dept_id: int = 1,
    recurse: bool = True,
    language: str = "zh_CN",
    page_size: int = 100,
    contain_access_limit: bool = False,
    fetch_detail: bool = False,
) -> List[JsonDict]:
    """获取通讯录全部用户（按 userid 去重）。

    Parameters
    ----------
    client:
        ``DingDiskClient``；默认 ``DingDiskClient.from_config()``。
    root_dept_id:
        起始部门，默认根部门 ``1``。
    recurse:
        ``True`` 时遍历整棵部门树；``False`` 仅拉 ``root_dept_id`` 直属成员。
    fetch_detail:
        ``True`` 时对每位用户再调 ``topapi/v2/user/get`` 补全详情
        （更慢，需成员信息读权限）。
    """
    if client is None:
        from api.ding_disk.client import DingDiskClient

        client = DingDiskClient.from_config()

    if recurse:
        departments = list_all_departments(
            client,
            root_dept_id=root_dept_id,
            language=language,
            include_root=True,
        )
        dept_ids = []
        for d in departments:
            try:
                dept_ids.append(int(d["dept_id"]))
            except (KeyError, TypeError, ValueError):
                continue
    else:
        dept_ids = [int(root_dept_id)]

    users_by_id: Dict[str, JsonDict] = {}
    for dept_id in dept_ids:
        if dept_id in _SKIP_DEPT_IDS:
            continue
        for user in iter_department_users(
            client,
            dept_id,
            page_size=page_size,
            language=language,
            contain_access_limit=contain_access_limit,
        ):
            userid = str(user.get("userid") or "").strip()
            if not userid:
                continue
            existing = users_by_id.get(userid)
            if existing is None:
                users_by_id[userid] = dict(user)
                continue
            # 合并所属部门
            merged_depts = _merge_dept_ids(
                existing.get("dept_id_list"),
                user.get("dept_id_list"),
                dept_id,
            )
            if merged_depts:
                existing["dept_id_list"] = merged_depts

    users = list(users_by_id.values())
    if fetch_detail:
        detailed: List[JsonDict] = []
        for user in users:
            userid = str(user.get("userid") or "").strip()
            try:
                detail = client.get_user(userid, language=language)
            except Exception:
                detailed.append(user)
                continue
            merged = dict(user)
            merged.update(detail)
            detailed.append(merged)
        users = detailed

    users.sort(key=lambda u: (str(u.get("name") or ""), str(u.get("userid") or "")))
    return users


def _merge_dept_ids(*sources: Any) -> List[int]:
    seen: Set[int] = set()
    out: List[int] = []
    for src in sources:
        if src is None:
            continue
        items: Iterable[Any]
        if isinstance(src, (list, tuple, set)):
            items = src
        else:
            items = [src]
        for item in items:
            try:
                dept_id = int(item)
            except (TypeError, ValueError):
                continue
            if dept_id in seen:
                continue
            seen.add(dept_id)
            out.append(dept_id)
    return out


def summarize_users(users: List[JsonDict]) -> JsonDict:
    """生成简要摘要，便于 CLI 预览。"""
    active = sum(1 for u in users if u.get("active") is True)
    admins = sum(1 for u in users if u.get("admin") is True)
    return {
        "total": len(users),
        "active": active,
        "admin": admins,
        "sample": [
            {
                "userid": u.get("userid"),
                "unionid": u.get("unionid"),
                "name": u.get("name"),
                "title": u.get("title"),
                "dept_id_list": u.get("dept_id_list"),
                "mobile": u.get("mobile"),
                "active": u.get("active"),
            }
            for u in users[:20]
        ],
    }


def _clip(value: Any, max_len: int) -> str:
    text = str(value or "").strip()
    if max_len > 0 and len(text) > max_len:
        return text[:max_len]
    return text


def _nullable(value: str) -> Optional[str]:
    """唯一键字段空串写 NULL，避免 UNIQUE 冲突。"""
    text = (value or "").strip()
    return text or None


def map_dingtalk_user(user: MappingLike) -> Optional[JsonDict]:
    """钉钉成员 → ``sys_user`` / ``sys_user_oauth`` 行字段。"""
    open_id = _clip(user.get("userid"), 128)
    if not open_id:
        return None
    name = _clip(user.get("name"), 64)
    email = _clip(user.get("email") or user.get("org_email"), 128)
    mobile = _clip(user.get("mobile"), 24)
    position = _clip(user.get("title"), 128)
    avatar = _clip(user.get("avatar"), 255)
    union_id = _clip(user.get("unionid") or user.get("unionId"), 128) or None
    active = user.get("active")
    status = 0 if active is False else 1
    username = _clip(f"ding:{open_id}", 50)
    return {
        "oauth_type": OAUTH_TYPE_DINGTALK,
        "open_id": open_id,
        "union_id": union_id,
        "nickname": name,
        "email": email,
        "mobile": mobile,
        "position": position,
        "oauth_nickname": _clip(name, 100),
        "oauth_avatar": avatar,
        "username": username,
        "phone": _clip(mobile, 20),
        "sys_email": _clip(email, 100),
        "avatar": avatar,
        "status": status,
    }


@dataclass
class SaveStats:
    total: int = 0
    skipped: int = 0
    oauth_inserted: int = 0
    oauth_updated: int = 0
    user_created: int = 0
    user_reused: int = 0
    user_updated: int = 0
    errors: List[str] = field(default_factory=list)

    def as_dict(self) -> JsonDict:
        return {
            "total": self.total,
            "skipped": self.skipped,
            "oauth_inserted": self.oauth_inserted,
            "oauth_updated": self.oauth_updated,
            "user_created": self.user_created,
            "user_reused": self.user_reused,
            "user_updated": self.user_updated,
            "errors": list(self.errors),
        }


def _find_oauth(cur: Any, open_id: str) -> Optional[JsonDict]:
    cur.execute(
        f"SELECT id, user_id FROM `{TABLE_OAUTH}` "
        f"WHERE oauth_type=%s AND open_id=%s LIMIT 1",
        (OAUTH_TYPE_DINGTALK, open_id),
    )
    row = cur.fetchone()
    return row if isinstance(row, dict) else None


def _find_sys_user(cur: Any, mapped: JsonDict) -> Optional[int]:
    """按 phone / email / username 匹配已有 ``sys_user``（未软删）。"""
    phone = _nullable(str(mapped.get("phone") or ""))
    email = _nullable(str(mapped.get("sys_email") or ""))
    username = _nullable(str(mapped.get("username") or ""))
    clauses: List[str] = []
    params: List[Any] = []
    if phone:
        clauses.append("phone=%s")
        params.append(phone)
    if email:
        clauses.append("email=%s")
        params.append(email)
    if username:
        clauses.append("username=%s")
        params.append(username)
    if not clauses:
        return None
    sql = (
        f"SELECT id FROM `{TABLE_USER}` "
        f"WHERE delete_time IS NULL AND ({' OR '.join(clauses)}) "
        f"ORDER BY id ASC LIMIT 1"
    )
    cur.execute(sql, tuple(params))
    row = cur.fetchone()
    if not row:
        return None
    return int(row["id"])


def _insert_sys_user(cur: Any, mapped: JsonDict) -> int:
    cur.execute(
        f"""
        INSERT INTO `{TABLE_USER}`
            (username, phone, email, password, salt, nickname, avatar, gender, status)
        VALUES
            (%s, %s, %s, NULL, NULL, %s, %s, 0, %s)
        """,
        (
            _nullable(str(mapped.get("username") or "")),
            _nullable(str(mapped.get("phone") or "")),
            _nullable(str(mapped.get("sys_email") or "")),
            mapped.get("nickname") or "",
            mapped.get("avatar") or "",
            int(mapped.get("status") if mapped.get("status") is not None else 1),
        ),
    )
    return int(cur.lastrowid)


def _update_sys_user(cur: Any, user_id: int, mapped: JsonDict) -> bool:
    """补全/刷新绑定用户基础资料；phone/email 仅在原值为空时写入。"""
    cur.execute(
        f"SELECT nickname, avatar, status, phone, email FROM `{TABLE_USER}` "
        f"WHERE id=%s LIMIT 1",
        (user_id,),
    )
    row = cur.fetchone()
    if not row:
        return False
    new_phone = _nullable(str(mapped.get("phone") or ""))
    new_email = _nullable(str(mapped.get("sys_email") or ""))
    phone = row.get("phone") or new_phone
    email = row.get("email") or new_email
    nickname = mapped.get("nickname") or row.get("nickname") or ""
    avatar = mapped.get("avatar") or row.get("avatar") or ""
    status = int(mapped.get("status") if mapped.get("status") is not None else row.get("status") or 1)
    cur.execute(
        f"""
        UPDATE `{TABLE_USER}`
        SET nickname=%s, avatar=%s, status=%s, phone=%s, email=%s
        WHERE id=%s
        """,
        (nickname, avatar, status, phone, email, user_id),
    )
    return cur.rowcount > 0


def _upsert_oauth(cur: Any, user_id: int, mapped: JsonDict, *, exists: bool) -> None:
    if exists:
        cur.execute(
            f"""
            UPDATE `{TABLE_OAUTH}`
            SET nickname=%s,
                email=%s,
                mobile=%s,
                position=%s,
                union_id=%s,
                oauth_nickname=%s,
                oauth_avatar=%s
            WHERE oauth_type=%s AND open_id=%s
            """,
            (
                mapped.get("nickname") or "",
                mapped.get("email") or "",
                mapped.get("mobile") or "",
                mapped.get("position") or "",
                mapped.get("union_id"),
                mapped.get("oauth_nickname") or "",
                mapped.get("oauth_avatar") or "",
                OAUTH_TYPE_DINGTALK,
                mapped["open_id"],
            ),
        )
        return
    cur.execute(
        f"""
        INSERT INTO `{TABLE_OAUTH}`
            (user_id, oauth_type, nickname, email, mobile, position,
             open_id, union_id, oauth_nickname, oauth_avatar)
        VALUES
            (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """,
        (
            user_id,
            OAUTH_TYPE_DINGTALK,
            mapped.get("nickname") or "",
            mapped.get("email") or "",
            mapped.get("mobile") or "",
            mapped.get("position") or "",
            mapped["open_id"],
            mapped.get("union_id"),
            mapped.get("oauth_nickname") or "",
            mapped.get("oauth_avatar") or "",
        ),
    )


def save_users_to_oauth(
    users: List[JsonDict],
    *,
    dry_run: bool = False,
    conn: Any = None,
) -> SaveStats:
    """将钉钉成员写入 ``sys_user_oauth``（必要时创建 ``sys_user``）。

    唯一键：``(oauth_type='dingtalk', open_id=userid)``。
    """
    import pymysql.cursors
    from database.db_connection import get_db_manager

    stats = SaveStats(total=len(users))
    own_conn = conn is None
    if own_conn:
        conn = get_db_manager().get_connection()

    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for user in users:
                mapped = map_dingtalk_user(user)
                if mapped is None:
                    stats.skipped += 1
                    continue
                open_id = mapped["open_id"]
                try:
                    existing = _find_oauth(cur, open_id)
                    if existing:
                        user_id = int(existing["user_id"])
                        stats.user_reused += 1
                        if dry_run:
                            stats.oauth_updated += 1
                            stats.user_updated += 1
                            continue
                        if _update_sys_user(cur, user_id, mapped):
                            stats.user_updated += 1
                        _upsert_oauth(cur, user_id, mapped, exists=True)
                        stats.oauth_updated += 1
                        continue

                    user_id = _find_sys_user(cur, mapped)
                    if user_id is None:
                        if dry_run:
                            stats.user_created += 1
                            stats.oauth_inserted += 1
                            continue
                        user_id = _insert_sys_user(cur, mapped)
                        stats.user_created += 1
                    else:
                        stats.user_reused += 1
                        if dry_run:
                            stats.oauth_inserted += 1
                            stats.user_updated += 1
                            continue
                        if _update_sys_user(cur, user_id, mapped):
                            stats.user_updated += 1

                    _upsert_oauth(cur, user_id, mapped, exists=False)
                    stats.oauth_inserted += 1
                except Exception as exc:
                    stats.errors.append(f"{open_id}: {exc}")
                    conn.rollback()
                    # 继续处理下一位；rollback 后由后续语句重新开事务
                    continue

            if dry_run:
                conn.rollback()
            else:
                conn.commit()
    finally:
        if own_conn:
            conn.close()
    return stats


def main(argv: Optional[List[str]] = None) -> int:
    _bootstrap()
    from api.ding_disk.client import DingDiskClient
    from api.ding_disk.exceptions import DingDiskError

    parser = argparse.ArgumentParser(description="获取钉钉通讯录全部用户信息")
    parser.add_argument(
        "--dept-id",
        type=int,
        default=1,
        help="起始部门 ID，默认根部门 1",
    )
    parser.add_argument(
        "--no-recurse",
        action="store_true",
        help="不递归子部门，仅拉取 --dept-id 直属成员",
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=100,
        help="部门用户分页大小，1～100，默认 100",
    )
    parser.add_argument(
        "--contain-access-limit",
        action="store_true",
        help="是否返回访问受限员工",
    )
    parser.add_argument(
        "--fetch-detail",
        action="store_true",
        help="再调 user/get 补全每位用户详情（更慢）",
    )
    parser.add_argument(
        "--list-depts",
        action="store_true",
        help="仅列出全部部门后退出",
    )
    parser.add_argument(
        "--summary",
        action="store_true",
        help="只打印摘要（人数 + 前 20 条精简字段）",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="将完整 JSON 写入文件（UTF-8）",
    )
    parser.add_argument(
        "--save",
        action="store_true",
        help="写入 sys_user / sys_user_oauth（oauth_type=dingtalk）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="配合 --save：只统计将执行的写入，不提交",
    )
    args = parser.parse_args(argv)

    try:
        client = DingDiskClient.from_config()

        if args.list_depts:
            depts = list_all_departments(
                client,
                root_dept_id=args.dept_id,
                include_root=True,
            )
            payload: Any = depts
            print(json.dumps(payload, ensure_ascii=False, indent=2))
            print(f"[OK] departments={len(depts)}", file=sys.stderr)
            return 0

        users = get_all_users(
            client,
            root_dept_id=args.dept_id,
            recurse=not args.no_recurse,
            page_size=args.page_size,
            contain_access_limit=args.contain_access_limit,
            fetch_detail=args.fetch_detail,
        )

        if args.save:
            stats = save_users_to_oauth(users, dry_run=args.dry_run)
            print(json.dumps(stats.as_dict(), ensure_ascii=False, indent=2))
            mode = "dry-run" if args.dry_run else "saved"
            print(
                f"[OK] users={len(users)} {mode} "
                f"oauth+={stats.oauth_inserted} oauth~={stats.oauth_updated} "
                f"errors={len(stats.errors)}",
                file=sys.stderr,
            )
            if args.out:
                out_path = Path(args.out)
                out_path.parent.mkdir(parents=True, exist_ok=True)
                out_path.write_text(
                    json.dumps(users, ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
            return 1 if stats.errors else 0

        payload = summarize_users(users) if args.summary else users
        text = json.dumps(payload, ensure_ascii=False, indent=2)
        if args.out:
            out_path = Path(args.out)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(text + "\n", encoding="utf-8")
            print(f"[OK] users={len(users)} written={out_path}", file=sys.stderr)
        else:
            print(text)
            print(f"[OK] users={len(users)}", file=sys.stderr)
        return 0
    except DingDiskError as exc:
        print(f"[ERR] {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"[ERR] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
