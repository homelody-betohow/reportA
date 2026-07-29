"""同步钉钉通讯录用户到 ``sys_user`` / ``sys_user_oauth``。

拉取逻辑与落库逻辑委托 ``api.ding_disk.getUsers``：
1. ``get_all_users``：递归部门树，按 userid 去重；
2. ``save_users_to_oauth``：按 ``(oauth_type=dingtalk, open_id=userid)`` upsert。

用法（项目根目录）::

    python app/ding-disk/syncUsers.py
    python app/ding-disk/syncUsers.py --dry-run
    python app/ding-disk/syncUsers.py --summary
    python app/ding-disk/syncUsers.py --dept-id 1 --no-recurse
    python app/ding-disk/syncUsers.py --fetch-detail --out users.json
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any, List, Optional

_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(_epr_mod)
_PROJECT_ROOT = _epr_mod.bootstrap(__file__)

from api.ding_disk.client import DingDiskClient  # noqa: E402
from api.ding_disk.exceptions import DingDiskError  # noqa: E402
from api.ding_disk.getUsers import (  # noqa: E402
    get_all_users,
    save_users_to_oauth,
    summarize_users,
)


def sync_users(
    *,
    root_dept_id: int = 1,
    recurse: bool = True,
    page_size: int = 100,
    contain_access_limit: bool = False,
    fetch_detail: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """拉取钉钉用户并写入 ``sys_user_oauth``，返回统计结果。"""
    client = DingDiskClient.from_config()
    users = get_all_users(
        client,
        root_dept_id=root_dept_id,
        recurse=recurse,
        page_size=page_size,
        contain_access_limit=contain_access_limit,
        fetch_detail=fetch_detail,
    )
    stats = save_users_to_oauth(users, dry_run=dry_run)
    result = stats.as_dict()
    result["fetched"] = len(users)
    result["dry_run"] = dry_run
    return result


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="同步钉钉通讯录用户到 sys_user / sys_user_oauth",
    )
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
        "--dry-run",
        action="store_true",
        help="只统计将执行的写入，不提交",
    )
    parser.add_argument(
        "--summary",
        action="store_true",
        help="同步前先打印拉取摘要（人数 + 前 20 条）",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="将拉取到的完整用户 JSON 写入文件（UTF-8）",
    )
    args = parser.parse_args(argv)

    try:
        client = DingDiskClient.from_config()
        users = get_all_users(
            client,
            root_dept_id=args.dept_id,
            recurse=not args.no_recurse,
            page_size=args.page_size,
            contain_access_limit=args.contain_access_limit,
            fetch_detail=args.fetch_detail,
        )
        print(f"[OK] fetched users={len(users)}", file=sys.stderr)

        if args.summary:
            print(json.dumps(summarize_users(users), ensure_ascii=False, indent=2))

        if args.out:
            out_path = Path(args.out)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(
                json.dumps(users, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            print(f"[OK] written={out_path}", file=sys.stderr)

        stats = save_users_to_oauth(users, dry_run=args.dry_run)
        print(json.dumps(stats.as_dict(), ensure_ascii=False, indent=2))
        mode = "dry-run" if args.dry_run else "saved"
        print(
            f"[OK] {mode} "
            f"oauth+={stats.oauth_inserted} oauth~={stats.oauth_updated} "
            f"user+={stats.user_created} user~={stats.user_updated} "
            f"errors={len(stats.errors)}",
            file=sys.stderr,
        )
        return 1 if stats.errors else 0
    except DingDiskError as exc:
        print(f"[ERR] {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"[ERR] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
