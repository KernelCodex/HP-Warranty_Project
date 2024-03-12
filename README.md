#

Create Virtual Environment for isolated dependencies

Create Environment
python -m venv .venv

Activate Environment
.venv\Scripts\activate

Deactivate Environment
deactivate

Execution Restrictions in Directory

Restricted Access
Set-ExecutionPolicy Unrestricted -Scope Process

Unrestricted Access
Set-ExecutionPolicy Restricted -Scope Process

Initializing Dependencies

Creating requirements
pip freeze > requirements.txt

Install dependencies
pip install -r requirements.txt
