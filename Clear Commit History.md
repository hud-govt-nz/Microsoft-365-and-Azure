
# Erase GitHub Repository Commit History

Follow these steps to erase the commit history of a GitHub repository. This process will remove all previous commits and start your repository with a new initial commit.

## Warning
This process is irreversible. Ensure you have a backup of your code or you are certain you want to remove all previous commits.

## Steps to Erase Commit History

1. **Clone the Repository**
   First, clone the repository to your local machine (if not already done).
   ```bash
   git clone https://github.com/your_username/your_repository.git
   cd your_repository
   ```

2. **Checkout**
   Checkout to the branch from which you want to remove the commit history.
   ```bash
   git checkout master # or any other branch
   ```

3. **Remove History**
   Create a fresh temporary branch and switch to it.
   ```bash
   git checkout --orphan temp_branch
   ```

4. **Add All Files**
   Add all the files to this temporary branch.
   ```bash
   git add -A
   ```

5. **Commit Changes**
   Commit the changes. This will be your new initial commit.
   ```bash
   git commit -am "Initial commit"
   ```

6. **Delete the Old Branch**
   Delete the old branch (e.g., master).
   ```bash
   git branch -D master
   ```

7. **Rename the Temporary Branch**
   Rename the temporary branch to master or your original branch name.
   ```bash
   git branch -m master
   ```

8. **Force Push to GitHub**
   Finally, force push the changes to GitHub. This will overwrite the old commit history.
   ```bash
   git push -f origin master
   ```

## Conclusion
Your GitHub repository now has only one commit. Ensure to inform any collaborators, as this change will affect everyone working on the project.
