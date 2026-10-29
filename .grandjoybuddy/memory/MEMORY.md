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
