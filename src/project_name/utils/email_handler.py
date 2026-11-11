# --------------------------- DOCUMENTATION ---------------------------#
# --------------------------- IMPORTS ---------------------------#
import os
import webbrowser
from typing import Any

from jinja2 import Template


# --------------------------- CONSTANTS ---------------------------#
# --------------------------- FUNCTIONS ---------------------------#
def display_draft_email(
    input_template_path: str,
    output_dir: str,
    file_name: str | None,
    to: list[str] | None = None,
    subject: str = "",
    body: str = "",
    cc: list[str] | None = None,
    **kwargs: Any,
):
    """Creates and displays HTML body to be copied to your email. Applicable to use since Outlook require admin acess\n
    **input_template_path:** html template file
    """
    with open(input_template_path, "r", encoding="utf-8") as f:
        body = Template(f.read()).render(to=to, cc=cc, sub=subject, **kwargs)

        if not file_name:
            file_name = "temp"
        temp_file = os.path.join(output_dir, f"{file_name}.html")
        with open(temp_file, "w", encoding="utf-8") as f:
            f.write(body)

        webbrowser.open(temp_file)

        # Remove email file
        # input("Press Enter after copying the email body...")
        # os.remove(temp_file)


# --------------------------- MAIN CODE ---------------------------#
if __name__ == "__main__":
    pass
