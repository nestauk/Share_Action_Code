# Share_Action_Code

## Setup

This project uses [`uv`](https://docs.astral.sh/uv/) for virtual environment management. If you are new to `uv`, you can find the [quickstart guide here](https://docs.astral.sh/uv/getting-started/).

We also utilise `direnv` via the `.envrc` file to automatically:

- Import your environment variables from `.env`
- Activate your virtual environment (_only if you comment out the relevant lines in `.envrc`_)

After installing `direnv` and `uv` on your system (we recommend doing this via [`brew`](https://brew.sh/) on macOS), set up the project:

- Clone the repo and navigate to your local repository folder.
- Run the following lines in your terminal to set up and activate your environment:

```bash
direnv allow
uv sync
uv run pre-commit install --install-hooks
source .venv/bin/activate
```
---
