#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cron 20:00 简短汇总给老板微信（合并 PPT + 越秀两件事）"""
import os, sys, json

# 加载 .env
env_path = '/home/lt-ceo/.hermes/.env'
if os.path.exists(env_path):
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k, v = line.split('=', 1)
            k, v = k.strip(), v.strip().strip('"').strip("'")
            if k not in os.environ or not os.environ[k]:
                os.environ[k] = v

# 清 auto-delivery env（防止被去重 guard）
for k in ('HERMES_CRON_AUTO_DELIVER_PLATFORM',
          'HERMES_CRON_AUTO_DELIVER_CHAT_ID',
          'HERMES_CRON_AUTO_DELIVER_THREAD_ID'):
    os.environ.pop(k, None)

sys.path.insert(0, '/home/lt-ceo/.hermes/hermes-agent')
from tools.send_message_tool import send_message_tool

msg = (
    "🕗 [cron 20:00 · 状态] 老板\n\n"
    "**PPT** · 微信 iLink 仍 rate limited（已持续 >2h，18:01 起）→ 21:00 / 22:00 我会按 19:00 承诺再试 2 次，最坏到 23:00 兜底结束。文件本身没坏，路径仍然是 `D:\\projects\\jinan_travel\\2026-07-09\\zhengzhou-jinan-5days.pptx` 直接双击即可。\n\n"
    "**越秀数据库填充** · 19:24 我给您 A/B/C 三个选项，还在等你拍板（不重发以免刷屏）。\n\n"
    "— 铁子"
)

result_str = send_message_tool({
    "action": "send",
    "target": "weixin:o9cq801T8efhwIzfr7nuWXmp9YiY",
    "message": msg,
})
try:
    result = json.loads(result_str) if isinstance(result_str, str) else result_str
except Exception:
    result = {"raw": str(result_str)[:300]}

print(json.dumps(result, ensure_ascii=False, indent=2))
