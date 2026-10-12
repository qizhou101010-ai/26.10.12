# 26.10.12 — Excel 填报数据核查 skill

一个 Claude Code skill：按 Excel 每行引用的《资料》到本地资料库查原文，核对数值、坐标、来源日期、名称和字符要求，输出错误报告和修正版 Excel；也能只给对象名称生成填写要求模板并按资料自动填表。程序优先，默认不调用大模型。

## 安装

把本仓库克隆到 Claude Code 的 skills 目录：

```bash
git clone https://github.com/qizhou101010-ai/26.10.12 ~/.claude/skills/excel-source-audit
```

第一次使用时运行 `python scripts/ensure_env.py` 建立 Python 环境（需要 Python 3.10+）。之后在 Claude Code 里直接说“帮我核一下待审核里的表”之类的话即可。

详细用法见 [SKILL.md](SKILL.md)。
