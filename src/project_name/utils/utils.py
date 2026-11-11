# --------------------------- DOCUMENTATION ---------------------------#
"""Contains small helper functions that don't belong to a specific class."""


# --------------------------- IMPORTS ---------------------------#
# --------------------------- CONSTANTS ---------------------------#
# --------------------------- FUNCTIONS ---------------------------#
def continue_process(input_text: str | list[str]):
    """Confirmation before proceeding to next step"""

    input_str = "\n".join(input_text) if type(input_text) == list else input_text
    continue_process = input(f"{input_str} \nType 'Y' to continue and 'N' to stop: ")
    if continue_process.strip().upper() == "N":
        raise ValueError("Process stopped by user")
    elif continue_process.strip().upper() != "Y":
        raise ValueError("Incorrect user input")
    else:
        return


# --------------------------- MAIN CODE ---------------------------#
if __name__ == "__main__":
    pass
