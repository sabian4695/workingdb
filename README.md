
# WorkingDB

*This is all of the files for WorkingDB, including all commands and function database. Some of the supporting documentation isn't here but that's about it.*

> The intention is for the GitHub Repository to be the most up-to-date production code. It also houses all of the in-process code, split into branches.

### Key Information:
- All .accdb/de files are stored with their decomposed files in a folder tree
- If you make a change, you must use a branch. Only a Code Owner can Pull the request and merge into the Main branch

## CONTRIBUTOR - To make a change to this repository you must:
1. Have a local version of this repository
2. Create a branch (in the remote repository)
3. Make sure you are on the branch in your local repository as well
4. Make your changes locally within the .accdb files
5. Decompose your changes into the files using the supplied .cmd file
6. Commit the changes using Git into the repository

Once this is complete, notify a Code Owner that the change is ready for review

## CODE OWNER - How accept/reject changes
1. Have a local version of this repository
2. Move to branch that is submitted for review
3. Review changes
4. Accept or Reject changes
5. Open a Pull request
	- Recompose accepted changed into .accdb file and clean DB if necessary
	- Revert other changes
6. Send feedback to contributor
7. Merge pull request with accepted changes in Main branch
