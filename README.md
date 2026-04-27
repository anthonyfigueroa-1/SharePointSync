# SharePoint Sync

SharePoint Sync is a Python script for syncing files from Microsoft SharePoint to either a PostgreSQL or SQLite database. Using OpenAI to help process photos into descriptive text, generating keywords for each document, and generating summaries for each document.

## Installation

You can use [uv](https://docs.astral.sh/uv/getting-started/installation/) to run the script after install.

```
git clone https://github.com/anthonyfigueroa-1/SharePointSync

cd SharePointSync

uv sync
```

After syncing, uv should install all module dependencies.

## Usage

To use you must define the below environment variables. The document containing these variables should be called ".env" and be saved in the root directory of the script. (Where main.py is.)

As of right now. Only .docx files are supported using this script.

```
example.env

OPENAI_KEY = <OpenAI API Key>
DOCX_INSTRUCTIONS = <Path to instructions for .docx files being processed by OpenAI>

PROCESS_FILETYPES = [".docx"]

DATABASE_URL = <PostgreSQL or SQLite database url>
SOLUTION_DOCS_TABLE_NAME = "sharepoint_solution_docs_table_example"
SOLUTION_DOCS_TABLE_SCHEMA = "public"
# "public" can be left as default unless you already have a seperate schema for the table to reside in.

TENANT_ID = <Tenant ID for your Microsoft Azure/Entra tenant>

SHAREPOINT_DOMAIN = <SharePoint site domain>
# Example, if epicdomain.sharepoint.com, then "epicdomain" should be set for this variable

SHAREPOINT_SITE_NAME = <Name of SharePoint site>
# Example, if epicdomain.sharepoint.com/site/TrashTake/..., then "TrashTake" should be set for this variable

SHAREPOINT_SITE_ROOT_DIR = "/root/children"
# This can more carefully defined if you want only a specific folder from the SharePoint site. Otherwise, leave as "/root/children"

TZ = "America/Los_Angeles"
# Change timezone as needed.
```

After defining the ENV variables. You will need to ensure that you have a Microsoft account with sufficient permissions to the SharePoint site defined in the .env.

In the root of the project, run...

```
uv run alembic revision --autogenerate -m "Initial Schema" && uv run alembic upgrade head
# The above command is to generate a script, using Alembic to be able to have Alembic create the table in the database url stated in the .env file. And the second command actually runs the script to create the table.

uv run main.py
```

And after a short while you should be prompted to sign in to Microsoft. After signing in, will the SharePoint files start to sync to the database of your choosing.
