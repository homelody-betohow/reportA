"""命令行冒烟：验证钉钉凭证、通讯录与文档表格接口。

先在 ``api/ding_disk/config.py`` 填写 ``APP_KEY`` / ``APP_SECRET`` / ``OPERATOR_ID``，
以及可选的 ``WORKBOOK_ID``，再执行::

    python -m api.ding_disk.smoke_test
    python -m api.ding_disk.smoke_test --token-only
    python -m api.ding_disk.smoke_test --user-id 016067253334-1323510411
    python -m api.ding_disk.smoke_test --mobile 13800138000
    python -m api.ding_disk.smoke_test --workbook-id YOUR_ID --list-sheets
    python -m api.ding_disk.smoke_test --workbook-id YOUR_ID --sheet Sheet1 --range A1:C3
    python -m api.ding_disk.smoke_test --workbook-id YOUR_ID --sheet Sheet1 --append '[["a","b"]]'
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path


def _bootstrap():
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def main(argv=None) -> int:
    _bootstrap()
    from api.ding_disk import DingDiskClient, workbook_id_from_url
    from api.ding_disk.config import DingDiskConfig
    from api.ding_disk.exceptions import DingDiskError

    parser = argparse.ArgumentParser(description="钉钉文档表格 / 通讯录 API 冒烟测试")
    parser.add_argument("--token-only", action="store_true", help="仅获取 accessToken")
    parser.add_argument("--user-id", default=None, help="通讯录：按 userId 查详情（含 unionId）")
    parser.add_argument("--mobile", default=None, help="通讯录：按手机号查 userId")
    parser.add_argument("--resolve-operator", action="store_true", help="解析 config.OPERATOR_ID → unionId")
    parser.add_argument("--workbook-id", default=None, help="表格 workbookId / nodeId")
    parser.add_argument("--workbook-url", default=None, help="从表格 URL 解析 workbookId")
    parser.add_argument("--list-sheets", action="store_true", help="列出全部工作表")
    parser.add_argument("--sheet", default=None, help="工作表 ID 或名称")
    parser.add_argument("--range", dest="range_address", default=None, help="A1 区域，如 A1:C3")
    parser.add_argument("--select", default="values", help="get_range 的 select，默认 values")
    parser.add_argument(
        "--append",
        default=None,
        help='追加行 JSON，如 [["col1","col2"],["1","2"]]',
    )
    parser.add_argument(
        "--write",
        default=None,
        help="写入起点单元格的 JSON 二维数组，配合 --start-cell",
    )
    parser.add_argument("--start-cell", default="A1", help="write 起点，默认 A1")
    parser.add_argument("--raw", action="store_true", help="打印完整 JSON")
    args = parser.parse_args(argv)

    try:
        cfg = DingDiskConfig.default()
        client = DingDiskClient(cfg)

        if args.token_only:
            token = client.get_access_token(force_refresh=True)
            print(f"[OK] accessToken acquired, length={len(token)}")
            if args.raw:
                print(token)
            return 0

        if args.mobile:
            userid = client.get_userid_by_mobile(args.mobile)
            print(json.dumps({"mobile": args.mobile, "userid": userid}, ensure_ascii=False))
            return 0

        if args.user_id:
            detail = client.get_user(args.user_id)
            print(json.dumps(detail, ensure_ascii=False, indent=2 if args.raw else None))
            return 0

        if args.resolve_operator:
            unionid = client.resolve_operator_union_id()
            print(
                json.dumps(
                    {"operatorId": cfg.operator_id, "unionId": unionid},
                    ensure_ascii=False,
                )
            )
            return 0

        workbook_id = args.workbook_id or cfg.workbook_id
        if args.workbook_url:
            parsed = workbook_id_from_url(args.workbook_url)
            if not parsed:
                print(f"[FAIL] 无法从 URL 解析 workbookId: {args.workbook_url}", file=sys.stderr)
                return 1
            workbook_id = parsed
            print(f"[OK] workbookId from url: {workbook_id}")

        if args.list_sheets:
            data = client.list_sheets(workbook_id)
            print(json.dumps(data, ensure_ascii=False, indent=2 if args.raw else None))
            return 0

        if args.append is not None:
            if not args.sheet:
                print("[FAIL] --append 需要同时指定 --sheet", file=sys.stderr)
                return 1
            values = json.loads(args.append)
            data = client.append_rows(args.sheet, values, workbook_id)
            print(json.dumps(data, ensure_ascii=False, indent=2 if args.raw else None))
            return 0

        if args.write is not None:
            if not args.sheet:
                print("[FAIL] --write 需要同时指定 --sheet", file=sys.stderr)
                return 1
            values = json.loads(args.write)
            data = client.write_values(
                args.sheet,
                values,
                workbook_id,
                start_cell=args.start_cell,
            )
            print(json.dumps(data, ensure_ascii=False, indent=2 if args.raw else None))
            return 0

        if args.range_address:
            if not args.sheet:
                print("[FAIL] --range 需要同时指定 --sheet", file=sys.stderr)
                return 1
            data = client.get_range(
                args.sheet,
                args.range_address,
                workbook_id,
                select=args.select,
            )
            print(json.dumps(data, ensure_ascii=False, indent=2 if args.raw else None))
            return 0

        if args.sheet:
            data = client.get_sheet(args.sheet, workbook_id)
            print(json.dumps(data, ensure_ascii=False, indent=2 if args.raw else None))
            return 0

        token = client.get_access_token()
        print(f"[OK] accessToken acquired, length={len(token)}")
        if cfg.operator_id:
            unionid = client.resolve_operator_union_id()
            print(f"[OK] operator userId/unionId={cfg.operator_id} → unionId={unionid}")
        if workbook_id:
            data = client.list_sheets(workbook_id)
            sheets = data.get("value") or []
            print(f"[OK] sheets count={len(sheets)}")
            print(json.dumps(data, ensure_ascii=False, indent=2 if args.raw else None))
        else:
            print(
                "[HINT] 未配置 WORKBOOK_ID。可用 --list-sheets --workbook-id xxx "
                "或在业务脚本中指定。"
            )
        return 0
    except DingDiskError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"[FAIL] {type(exc).__name__}: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
