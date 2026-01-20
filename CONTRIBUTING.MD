# CONTRIBUTOR - How to contribute to a WorkingDB repository:
*This is the full process to contribute.*
*Please read all instructions*

| # | First Time Setup Steps |
| ----------- | ----------- |
| 1 | Drive Setup |
| 2 | Gain GitHub Access | 
| 3 | Install and Set Up GIT | 
| 4 | Clone Repositories | 

### 1. Drive Setup
- Map Prod Location of the database
- For Strategic Planning, use this: "\\\data\mdbdata\WorkingDB\Strategic_Planning_FE\" to Drive Letter: B
  
### 2. Gain [GitHub](https://github.com/) Access
- Log in/Sign Up
- Send a Code Owner your username to gain access to repository
   
### 3. Install and set up [GIT](https://git-scm.com/install/windows)
- Download/Install
- Log in using GitHub
- Right click in the parent folder (typically "H:\dev\")
- Open GIT GUI
- Clone GIT Repositories
   
### Clone Repositories
- First, clone Code Review : https://github.com/vbadecoded/ms-access-code-review-git-wrapper
- Then, clone the repository you want to work on: i.e. Strategic Planning : https://github.com/workingdb/strategic_planning
- *Name folder to match the repository name - i.e. "strategic_planning"*

---
| # | Steps to Publish New Changes |
| ----------- | ----------- |
| 1 | GIT Process |
| 2 | Accept/Reject the Changes | 
| 3 | Release Changes to Production Repository | 

### GIT Process
1. Open Code Review Database
2. Select Repository to work on
3. Click "Status" to check status of repository changes
   This is the same as the Bash comand:
	```bash
	git status
	```
4. Click "Enable Shift" to allow using Shift Bypass on Database
5. Shift + Click "Open Database" to bypass startup procedures
6. Do your work on the MS Access Database
7. Click Clean Database
8. Shift + Click **Decompose**
9. Click Status to see what files were changed

#### Accept/Reject the Changes
1. Review changes using the Git GUI program or the Diff button
   This is the same as the Bash comands:
	```bash
	git diff
 	git gui
	```
2. Commit / Push to your branch (NOT Master)
   This is the same as the Bash comand:
	```bash
	git commit
 	git push origin branch-name
	```

#### Release Changes to Production Repository
1. AFTER PUSH, switch to master branch
   This is the same as the Bash comand:
	```bash
	git checkout master
	```
2. Merge your branch
   This is the same as the Bash comand:
	```bash
	git merge branch-name
	```
3. "Push" to make changes public in GitHub
      This is the same as the Bash comand:
	```bash
	git push origin master
	```
4. Publish (Pull) down in production location to actually change the production front end
   This is the same as the Bash comand (in production repository):
	```bash
	git pull origin master
	```

# CODE OWNER - How to accept/reject changes
1. Have a local version of this repository
2. Move to branch that is submitted for review
3. Review changes
4. Accept or Reject changes
5. Open a Pull request
	- Recompose accepted changed into .accdb file and clean DB if necessary
	- Revert other changes
6. Send feedback to contributor
7. Merge pull request with accepted changes in Main branch