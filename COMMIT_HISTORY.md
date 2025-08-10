# Commit History

d03afc6 - James Cotrotsios, 5 seconds ago : refactor: Move JsonIntegration to root
48842db - James Cotrotsios, 6 minutes ago : DOC: Update commit history
1ed0244 - James Cotrotsios, 8 minutes ago : Added JSON Integration Capabilities - v2
a01475d - James Cotrotsios, 22 minutes ago : DOC: Add commit history and rollback guide
88e9bc0 - James Cotrotsios, 23 minutes ago : Added JSON Conversion Capabilities - v1
803105a - James Cotrotsios, 20 hours ago : Initial commit

# How to Rollback to a Previous Commit

If the application becomes unstable, you can roll back to a previous version using the following steps. This process involves creating a new branch from a specific commit, which preserves the project's history while allowing you to revert the main branch.

## 1. Find the Commit Hash

First, you need to find the hash of the commit you want to revert to. You can find this in the commit history above.

## 2. Create a New Branch from the Commit

Create a new branch from the desired commit. This will create a new branch with the state of the project at that commit.

```bash
git checkout -b <new-branch-name> <commit-hash>
```

For example, to revert to commit `803105a`, you would run:

```bash
git checkout -b hotfix-revert 803105a
```

## 3. Push the New Branch to the Remote

Push the new branch to the remote repository.

```bash
git push origin <new-branch-name>
```

## 4. (Optional) Reset the Main Branch

If you want to completely revert the main branch to the state of the new branch, you can do a hard reset. **Warning: This will permanently delete any changes made after the commit you are reverting to.**

```bash
git checkout master
git reset --hard <commit-hash>
git push origin master --force
```
