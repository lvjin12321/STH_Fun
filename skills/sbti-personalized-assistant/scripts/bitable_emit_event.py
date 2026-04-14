#!/usr/bin/env python3
"""Emit telemetry events to a Lark Bitable table (Webhook version).
"""

from __future__ import annotations

import argparse
import json
import requests

# Webhook URLs
ENTRANCE_URL = "https://bytedance.larkoffice.com/base/workflow/webhook/event/UBYeaJIQNwvQidh1eSCcyZ6tn3e"
EXIT_URL = "https://bytedance.larkoffice.com/base/workflow/webhook/event/E2q8aCbv2w8NdehXb2ycy5ednyg"

def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--trigger_type", required=True, help='"安装开启" or "完成测试"')
    ap.add_argument("--result_code", default="")
    ap.add_argument("--result_name", default="")
    
    # Ignore extra arguments for backward compatibility
    ap.add_argument("--username", default="")
    ap.add_argument("--base_url", default="")
    ap.add_argument("--table_name", default="")
    ap.add_argument("--emit_time_ms", default=0, type=int)
    
    args = ap.parse_args()

    if args.trigger_type == "安装开启":
        webhook_url = ENTRANCE_URL
        payload = {
            "event_type": args.trigger_type
        }
    elif args.trigger_type == "完成测试":
        webhook_url = EXIT_URL
        payload = {
            "event_type": args.trigger_type,
            "result_code": args.result_code,
            "result_name": args.result_name
        }
    else:
        # Default to entrance or handle as error? 
        # User specified only these two, so we'll just use entrance as fallback if needed but be strict.
        webhook_url = ENTRANCE_URL
        payload = {
            "event_type": args.trigger_type
        }

    try:
        resp = requests.post(webhook_url, json=payload, timeout=10)
        resp.raise_for_status()
        
        print(json.dumps({
            "ok": True, 
            "status_code": resp.status_code,
            "message": f"Event '{args.trigger_type}' emitted successfully"
        }, ensure_ascii=False))
        
    except Exception as e:
        print(json.dumps({
            "ok": False, 
            "error": str(e)
        }, ensure_ascii=False))

if __name__ == "__main__":
    main()
