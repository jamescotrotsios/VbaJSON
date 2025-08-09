# Guide: Setting Up a New Project with a Local and Remote Git Repository

This guide will walk you through the steps to start a new project in Visual Studio Code, create a local Git repository, and link it to a new remote repository on GitHub.

## 1. Starting a New Project in Visual Studio Code

1.  **Open a new folder:**
    - Open Visual Studio Code.
    - Go to `File > Open Folder...` (or `File > Open...` on macOS).
    - Navigate to where you want to create your project, create a new folder, and open it. This folder will be the root of your project.

## 2. Creating a Local Git Repository

A local repository is a `.git` directory inside your project folder, where Git stores all the version control information for your project.

1.  **Open the Integrated Terminal:**

    - In Visual Studio Code, open the integrated terminal by going to `View > Terminal` or by using the shortcut `` ` `` (backtick).

2.  **Initialize the Repository:**
    - In the terminal, make sure you are in your project's root folder.
    - Run the following command to initialize a new Git repository:
      ```bash
      git init
      ```
    - This command creates a hidden `.git` subfolder in your project directory, which contains all the necessary repository files.

## 3. Creating a Remote Repository on GitHub

A remote repository is a version of your project that is hosted on the internet, in this case, on GitHub.

1.  **Using the GitHub CLI (Command Line Interface):**

    - Make sure you have the [GitHub CLI](https://cli.github.com/) installed and authenticated (`gh auth login`).
    - In your terminal, run the following command to create a new repository on GitHub. Replace `repo-name` with the desired name for your repository.
      ```bash
      gh repo create VbaJSON --public --source=. --remote=origin
      ```
      - `--public`: Creates a public repository. Use `--private` for a private one.
      - `--source=.`: Specifies that the current directory should be the source for the new repository.
      - `--remote=origin`: Sets the name of the remote to `origin`.

2.  **Alternative: Using the GitHub Website:**
    - Go to [github.com](https://github.com) and log in.
    - Click the `+` icon in the top-right corner and select `New repository`.
    - Give your repository a name, choose whether it should be public or private, and click `Create repository`.
    - GitHub will provide you with the commands to link your local repository to this new remote repository.

## 4. Linking the Local and Remote Repositories

If you didn't use the `gh repo create` command with the `--source` flag, you'll need to manually link your local repository to the remote one.

1.  **Add the Remote:**

    - If you created the repository on the GitHub website, copy the repository URL.
    - In your terminal, run the following command, replacing `<repository-url>` with the URL you copied:
      ```bash
      git remote add origin <repository-url>
      ```
      - This command tells your local Git repository where the remote version is located.

2.  **Verify the Remote:**
    - To check that the remote was added correctly, run:
      ```bash
      git remote -v
      ```
    - You should see the URL for `origin` listed for both fetch and push.

## 5. Making Your First Commit and Pushing to GitHub

1.  **Stage Your Files:**

    - Add the files you want to track to the staging area. To add all files, run:
      ```bash
      git add .
      ```

2.  **Commit Your Changes:**

    - Commit the staged files to your local repository with a descriptive message:
      ```bash
      git commit -m "Initial commit"
      ```

3.  **Push to GitHub:**
    - Push your local commits to the remote repository on GitHub:
      ```bash
      git push -u origin master
      ```
      - The `-u` flag sets the upstream branch, so in the future, you can just run `git push`.
      - Note: The default branch name might be `main` instead of `master`. Use `git branch` to check and adjust the command accordingly (`git push -u origin main`).

Your project is now set up with both a local and a remote Git repository!
