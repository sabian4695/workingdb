
# WorkingDB

*This is all of the files for WorkingDB, including all commands and function database. Some of the supporting documentation isn't here but that's about it.*

> The intention is for the GitHub Repository to be the most up-to-date production code. It also houses all of the in-process code, split into branches.

## Table of Contents
- [Key Information](#Key)
- [Contributor Information](#CONTRIBUTOR)
- [Code Owner Information](#CODE)

# Key Information:
- All .accdb/de files are stored with their decomposed files in a folder tree
- If you make a change, you must use a branch. Only a Code Owner can Pull the request and merge into the Main branch

## Structure of Repository
/workingdb
├── /Forms                      All decomposed forms for this .accdb file
│   └── /SubForms               All decomposed sub-forms from this .accdb file
├── /Modules                    All decomposed VBA modules for this .accdb file
├── /Queries                    All decomposed queries for this .accdb file
│   └── /SubQueries             All decomposed sub-queries from this .accdb file
├── /Reports                    All decomposed reports for this .accdb file
│   └── /SubReports             All decomposed sub-reports from this .accdb file
├── WorkingDB_FE.accdb          Master .accdb file
└── README.md

# CONTRIBUTOR - How to contribute to this repository:
> Here is the full setup to contribute to any WorkingDB database
> Please read all instructions
### Drive Setup
- Map Prod Location of the database:
- For Strategic Planning, use this: "\\\data\mdbdata\WorkingDB\Strategic_Planning_FE\" to Drive Letter: B
### [GitHub](https://github.com/)
1. Log in/Sign Up
2. Send a Code Owner your username to gain access to repository
### [GIT](https://git-scm.com/install/windows)
1. Download/Install
2. Log in using GitHub
3. Right click in the parent folder (typically "H:\dev\")
4. Open GIT GUI
5. Clone GIT Repositories
### Cloning Repositories
1. First, clone [Code Review](https://github.com/vbadecoded/ms-access-code-review-git-wrapper)
2. Then, clone the repository you want to work on: i.e. [Strategic Planning](https://github.com/workingdb/strategic_planning)
> *Name folder "strategic_planning"*

### GIT Process
1. Open Code Review Database
2. Select Repository to work on
3. Click "Status" to check status of repository changes
4. Click "Enable Shift" to allow using Shift Bypass on Database
5. Shift + Click "Open Database" to bypass startup procedures
6. Do your work on the MS Access Database
7. Click Clean Database
8. Shift + Click **Decompose**
9. Click Status to see what files were changed

#### Accept/reject your own changes:
1. Git DIFF / GUI
2. Commit / Push to your branch (NOT Master)

#### To Release
1. AFTER PUSH, switch to master branch
2. Merge your branch
3. "Push" to make changes public in GitHub
4. Publish (Pull) down in production location to actually change the production front end

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
