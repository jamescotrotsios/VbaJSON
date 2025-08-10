# Git Workflow Tutorial for the VbaJSON Repository

This tutorial provides a basic guide to using Git with this repository. Following these steps will help keep the project history clean and understandable.

## 1. Start Your Work with `git pull`

Before you start making any changes, it's crucial to make sure you have the latest version of the code. Run the following command to pull any recent changes from the remote repository:

```bash
git pull origin master
```

This prevents you from accidentally working on an outdated version of the code and reduces the chance of merge conflicts.

## 2. Make Your Changes

Now you can work on the files in the project. Add new files, edit existing ones, and build your features.

## 3. Stage Your Changes with `git add`

Once you have finished a piece of work (e.g., a new feature, a bug fix), you need to tell Git which files you want to include in your next commit. You can do this with the `git add` command.

To add a specific file:

```bash
git add path/to/your/file.bas
```

To add all new and modified files:

```bash
git add .
```

## 4. Commit Your Changes with `git commit`

A commit is a snapshot of your staged changes at a specific point in time. Each commit has a message that should briefly describe the changes you made.

```bash
git commit -m "Your descriptive commit message here"
```

**Commit Message Best Practices:**

- Keep it concise and descriptive.
- Use the present tense (e.g., "Add feature" not "Added feature").
- If it's a bug fix, you can prefix it with "fix:".
- If it's a new feature, you can prefix it with "feat:".
- For documentation changes, use "docs:".

## 5. Push Your Changes with `git push`

After committing your changes, you need to push them to the remote repository on GitHub so that others can see them.

```bash
git push origin master
```

## The Role of `COMMIT_HISTORY.md`

The `COMMIT_HISTORY.md` file in this repository serves as a human-readable log of all the important changes that have been made. After you push a new commit, you should also update this file.

**How to Update `COMMIT_HISTORY.md`:**

1.  Get the latest commit history with this command:
    ```bash
    git log --pretty=format:"%h - %an, %ar : %s"
    ```
2.  Copy the new commit line(s) from the output.
3.  Open `COMMIT_HISTORY.md` and paste the new line(s) at the top of the commit list.
4.  Save the file, then `add`, `commit`, and `push` the change to `COMMIT_HISTORY.md`.

This file provides a quick and easy way for everyone on the team to see the project's progress without having to use the `git log` command. It also contains instructions on how to roll back to a previous commit if something goes wrong.
