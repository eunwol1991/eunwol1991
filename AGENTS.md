# 协作开发规范（AGENTS.md）

## 1. 修改/Debug前，务必先拉取最新代码

- 在对已存在的文件进行修改或Debug前，必须先执行 `git pull`，确保本地代码与远程仓库保持同步。
- 这样可以最大限度减少 `merge conflict`（合并冲突）的发生概率。

## 2. Commit和Push前再次确认

- 在准备提交（commit）和推送（push）前，再次执行 `git pull`，防止期间有其他成员更新了代码。

## 3. 冲突处理原则

- 若出现合并冲突，开发者需**主动解决冲突**，确认无误后再commit和push。
- 必要时可@相关开发者协助。

## 4. Commit信息要求

- 提交信息需简洁明了，包含本次变更的核心内容，例如“fix: 修复数据重复问题”，“feat: 新增自动导出功能”等。

## 5. 其他说明（可选）
- 建议每次push前`git status`检查变动文件。
- 如需新建分支开发，请提前沟通。

---

## 参考示例流程

```bash
git pull origin main         # 拉取最新
# debug & 修改代码
git add .
git commit -m "fix: 修复xxx问题"
git pull origin main         # 再拉一次，防止有新变化
# 如有冲突，解决后继续
git push origin main         # 推送到远程
