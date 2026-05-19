# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an **AI招聘审美官** (AI Recruitment Judge) system - an HR automation system for screening and evaluating job candidates across three position categories:

- **交付 (Delivery)**: Technical delivery/project roles
- **研发 (Development/R&D)**: Software development roles
- **销售 (Sales)**: Sales positions

The system uses black-white-gray list rules (黑白灰清单) for candidate classification and quantitative scoring across multiple dimensions (底色/三板斧/AI能力).

## Core Skills Pipeline

Each position type has its own skill chain:

| Position | Resume Screening | Interview Evaluation | Excel Filling |
|----------|-----------------|---------------------|----------------|
| 交付 | `del-cr` | `del-IID` | `del-excel-filler` |
| 研发 | `dev-cr` | `dev-IID` | - |
| 销售 | `sale-cr` (cr-new) | `sale-IID` | `sale-excel-filler-local` |

**Unified Archive**: `candidate-records` manages all candidate records by phone number.

## Skill Trigger Rules

- Resume file submitted + "筛选/分析/评估" → Auto-trigger corresponding `*-cr` skill
- Resume result + Interview transcript + "面试评估/补全评分/综合评价" → Auto-trigger corresponding `*-IID` skill
- Resume/Interview text + "归档/存储/保存" → Auto-trigger `candidate-records`

## Key Skill Files

Each skill in `.claude/skills/{skill-name}/` contains:
- `SKILL.md` - Main skill definition
- `references/*.md` - Black-white-gray rules and scoring rubrics
- `scripts/fill_excel.py` - Excel filling script (for excel-filler skills)
- `template.xlsx` - Excel template

## Data Structure

```
record/
└── {电话号码}/
    ├── 简历/
    │   ├── 原始简历.{扩展名}
    │   └── 简历分析报告.md
    └── 面试/
        ├── 面试第{N}轮_对话.jsonl
        ├── 面试第{N}轮_原始文本.txt
        └── 面试第{N}轮_分析报告.md
```

## Common Commands

```bash
# Fill Excel (for delivery positions)
python .claude/skills/del-excel-filler/scripts/fill_excel.py data.json 初试

# Fill Excel (for sales positions)
python .claude/skills/sale-excel-filler-local/scripts/fill_excel.py data.json output.xlsx

# Convert interview transcript to JSONL
python .claude/skills/candidate-records/references/jsonl-parser.py input.txt output.jsonl
```

## Scoring Framework

Each position type has its own scoring dimensions:

**Delivery (交付)**:
- 底色: AI能力、兴趣、情商、狼性、智商
- 三板斧: 快、收款、扩项
- 单元三板斧: 画施工图、选好骨干、提高人效

**Development (研发)**:
- 底色: 兴趣、狼性、智商
- 三板斧: 打深井、简化、质量好
- AI能力

**Sales (销售)**:
- 底色: 兴趣、狼性、情商、搞成铁杆
- 三板斧: 销售能力、客户资源、团队管理 (for A7-A9/M1-M3)

Black-white-gray classification:
- **黑清单**: Direct rejection
- **灰清单**: Pending review
- **白清单**: Priority recommendation

## Settings

Claude settings are in `.claude/settings.json` (global) and `.claude/settings.local.json` (local overrides).