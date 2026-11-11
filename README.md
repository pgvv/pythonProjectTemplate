# README
## Steps
- Inital Setup
```bash
# Clone the repository
git clone https://github.com/pgvv/pythonProjectTemplate
cd pythonProjectTemplate
Remove-Item -Recurse -Force .git
```
- Delete src/old folder
- Delete src/project_name folder to your project
```bash
# Initialize the project
uv init
git remote add origin <new-url>
```
- Delete unrequired files and classes/functions
