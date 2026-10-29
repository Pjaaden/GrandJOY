# GrandJOY 项目长期笔记

## Git 同步
- 远端：`git@github.com:Pjaaden/GrandJOY.git`（SSH，账号 Pjaaden）
- 本机密钥 `~/.ssh/id_ed25519`（jaden.yaan@gmail.com）已绑定 GitHub，读写权限就绪
- 同步纪律：开工 `git pull --rebase origin main`；收工 `git add -A && git commit && git pull --rebase && git push`
- 凭证只放 `TRAE_LOCAL/`（.gitignore 已忽略）

## 任务文件夹约定（重要）
- Buddy 承接的每个任务，产出统一放 `04-buddy任务/buddy-YYYYMMDD-<任务简述>/`
- 每任务一个独立子文件夹，不与 00-03 业务目录（索引/展会/品牌/归档）混放
- 任务结束如为长期有效产出，登记到 `00-索引/项目文件索引.md`

## 项目定位
大悦城总部招商（运动户外 / 潮流服饰）品牌资源挖掘，与 Brand-Engine、Stock-Research 为三个独立项目/独立 git 仓库。

## 战略台账（长期资产，2026-10-29 建）
- 位置：`04-buddy任务/buddy-20261029-战略台账/`；飞书：我的空间 / 大悦商业 / 战略台账 /「全国运动男装战略品牌合同台账」
- 飞书表 token `JDKGs49LKhZRL9ts7mscg14inbg`；大悦商业 folder `WhB2fPI8kliF0cdoTr4cLdjBn2f`；战略台账 folder `ZRbRfWrV7lMEKgdYWPocyboBngf`
- 更新入口：`bash 04-buddy任务/buddy-20261029-战略台账/脚本/更新台账.sh`（新版 BI xlsx 放进 `源数据/`）
- **BI 明细粒度 = 合同期次**，面积类指标只在合同级看（期次累加会重复计数）
- 月面积折算口径 **365/12 = 30.4167 天**（不是 30 天）
- 飞书 CLI 在 `~/.grandjoybuddy/binaries/node/workspace/node_modules/.bin/lark-cli`（不在 PATH）

## 飞书 CLI 使用要点（踩过的坑）
- `+cells-set --cells` 空位必须用 `{}`，不能用 `{"value": null}` —— 否则 schema 校验失败且静默返回空
- `+cells-batch-clear --ranges` 用子表**显示名**（`Sheet名!A1:H200`），不接受 `--sheet-id`
- 覆盖刷新前先清空目标区域，避免行数变少残留旧行
- 表格数据写入用 `+table-put --sheets` typed 协议保类型（日期/金额/百分比）

