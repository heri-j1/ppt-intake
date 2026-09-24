# ppt-intake — PPT 材料整合 Skill

把**别人的 PPT 材料**（格式/模板/字体与你的不一致）**摄入**你自己的基准 PPT，
产出一版格式与基准完全一致的成品。核心思想：**基准 PPT 的每一页都是模板**——
新页 = 克隆基准同角色页 + 替换文本；绝不从零建页，绝不搬运材料页。

标准 **Agent Skill** 格式（SKILL.md + YAML frontmatter），支持 Agent Skills
规范的 agent（ZCode、Claude Code 等）即可使用。

## 快速开始

```bash
git clone https://github.com/heri-j1/ppt-intake && cd ppt-intake
pip install python-pptx        # 唯一依赖
./install.sh claude-user       # 装到 Claude Code（其他目标见下表）
```

然后对 agent 说一句话：

- 「把这个材料 PPT 的内容整合进我的 base.pptx，格式以我的为准」
- 「对比一下这两份 PPT 的结构差异」（只读模式，不改文件）

### 安装目标

| 目标 | 落点 | 生效范围 |
|---|---|---|
| `claude-user` | `~/.claude/skills/ppt-intake` | Claude Code 所有项目 |
| `claude-project` | `<本项目>/.claude/skills/ppt-intake` | Claude Code 当前项目 |
| `zcode-user` | `~/.zcode/skills/ppt-intake` | ZCode 所有项目 |
| `zcode-project` | `<本项目>/.zcode/skills/ppt-intake` | ZCode 当前项目 |
| `agents-user` / `agents-project` | `.agents/skills/` | 通用 AGENTS 生态（ZCode 亦识别） |

`./install.sh` 不带参数列出全部目标，`all` 一次全装。
不用安装器时，手动复制等价：`cp -R skills/ppt-intake ~/.claude/skills/`。

> 本工作区内 `.zcode/skills/ppt-intake` 是指向真源的 NTFS junction（开发时免双份维护），
> 安装器检测到链接会自动跳过，不会误删。

## 手动跑脚本

仓库内可用相对路径；agent 场景请遵循 SKILL.md 的 `${SKILL_DIR}` 硬规则，
用 skill 目录的绝对路径调用：

```bash
python skills/ppt-intake/scripts/inspect_pptx.py    <file.pptx> -o out.json   # 盘点
python skills/ppt-intake/scripts/merge_apply.py     base.pptx plan.json -o out.pptx [-m 材料目录]
python skills/ppt-intake/scripts/validate_output.py out.pptx --base base.pptx --plan plan.json
```

merge_apply 成功后会在输出旁写 `<out>.receipt.json` **逐页处置回执**
（untouched / updated / cloned / deleted），validate 自动探测并与计划精确对账，
让「未涉及的页保持原样」可验证。

## Demo（自带样例，可复现）

```bash
python samples/make_samples.py     # 生成格式冲突样例（16:9蓝 vs 4:3红）
# 按 tests/smoke-test.md 的 T1–T6 执行
```

已验证产物：`samples/base_merged.pptx`（11 页）+ `samples/base_merged.receipt.json`
（处置回执）+ `samples/变更说明.md`，验证 10 项全绿（含回执对账与替换落地检查）：
克隆页字体继承基准（微软雅黑），材料宋体零泄漏，
冲突数据（竞品 A 售价 99 vs 89 元）按规则留待用户裁决。

## 结构

```
skills/ppt-intake/            # skill 本体（唯一真源，安装即复制它）
├── SKILL.md                  # 路由中心（agent 入口）
├── rules/                    # 铁律 + 映射决策规则
├── workflows/                # 主工作流 / 只读工作流
├── references/               # pptx 内部结构 / 坑集
├── scripts/                  # inspect / merge_apply / validate
└── templates/merge_plan.md   # 合并计划 JSON schema
install.sh                    # 安装器
samples/                      # 端到端样例 + 已验证产物
tests/smoke-test.md           # 全链路冒烟
.zcode/                       # ZCode 工作区配置（plans 存档 + 指向 skill 的 junction）
```

## 可选：接入本地 Qwen3-8B（离线模式）

skill 的 LLM 环节由宿主 agent 的模型承担。若要离线/私有化，可让本地模型起草
映射表初稿，agent 只做复核：

```bash
ollama pull qwen3:8b          # ≈5GB，建议 8GB+ 显存（纯 CPU 亦可但较慢）
ollama serve                  # OpenAI 兼容接口 http://localhost:11434/v1
```

```python
from openai import OpenAI
client = OpenAI(base_url="http://localhost:11434/v1", api_key="ollama")
resp = client.chat.completions.create(
    model="qwen3:8b",
    messages=[{"role": "user", "content": f"按以下规则把材料页映射到基准页...{prompt}"}],
    response_format={"type": "json_object"},   # Ollama 支持结构化输出
)
```

注意：8B 模型适合**粗映射 + 摘要改写**；冲突裁决与最终文案仍建议由 agent/人复核
（本 skill 的 conflict 机制正是为此设计的兜底）。

## 设计来源与致谢

- 工作流方法论：[linux.do 商业 PPT Agent 思路分享](https://linux.do/t/topic/1782304)（调研→大纲→策划→设计的分层思想）
- Skill 组织形式：[linux.do 如何写一个好的 Skill](https://linux.do/t/topic/1923706)（路由式 SKILL.md / rules / workflows / references 分层）
- 克隆/替换/预算规则：参考 Anthropic pptx skill 与 PPTAgent 的 template-inheritance 实践
- 「未涉及页保持不变」「确认后再执行」契约：借鉴 [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master)
