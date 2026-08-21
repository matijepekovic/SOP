from __future__ import annotations

from dataclasses import dataclass, field

from sop_reporter.exceptions import CredentialsCancelledError, CredentialsError


@dataclass(frozen=True)
class EnteredCredentials:
    account: str
    app_password: str = field(repr=False)


def prompt_for_credentials(
    prefilled_account: str = "", problem: str = ""
) -> EnteredCredentials:
    """Ask for Gmail credentials.

    ``problem`` reports why a previous attempt was refused, so a re-prompt
    explains itself instead of appearing again for no visible reason.
    """
    import tkinter as tk
    from tkinter import messagebox, ttk

    result: EnteredCredentials | None = None

    root = tk.Tk()
    root.title("SOP Reporter - Gmail Setup")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    frame = ttk.Frame(root, padding=20)
    frame.grid(row=0, column=0, sticky="nsew")

    ttk.Label(
        frame,
        text="Connect SOP Reporter to Gmail",
        font=("Segoe UI", 14, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
    ttk.Label(
        frame,
        text=(
            "Enter the Gmail address that receives the reports and a Google\n"
            "app password. The app password is stored in Windows Credential\n"
            "Manager and is never written to the configuration file."
        ),
        justify="left",
    ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 16))

    if problem:
        ttk.Label(
            frame,
            text=problem,
            justify="left",
            wraplength=380,
            foreground="#B00020",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(12, 0))

    account_var = tk.StringVar(value=prefilled_account)
    password_var = tk.StringVar()

    ttk.Label(frame, text="Gmail address").grid(row=2, column=0, sticky="w", pady=4)
    account_entry = ttk.Entry(frame, textvariable=account_var, width=42)
    account_entry.grid(row=2, column=1, sticky="ew", pady=4)

    ttk.Label(frame, text="App password").grid(row=3, column=0, sticky="w", pady=4)
    password_entry = ttk.Entry(frame, textvariable=password_var, width=42, show="*")
    password_entry.grid(row=3, column=1, sticky="ew", pady=4)

    def submit() -> None:
        nonlocal result
        account = account_var.get().strip()
        password = "".join(password_var.get().split())
        if not account or "@" not in account:
            messagebox.showerror("Invalid Gmail address", "Enter a complete email address.")
            return
        if not password:
            messagebox.showerror("Missing app password", "Enter the Gmail app password.")
            return
        if len(password) < 12:
            messagebox.showerror(
                "Invalid app password",
                "The app password is too short. Use the app password from Google, not the regular Gmail password.",
            )
            return
        result = EnteredCredentials(account=account, app_password=password)
        root.destroy()

    def cancel() -> None:
        root.destroy()

    button_frame = ttk.Frame(frame)
    button_frame.grid(row=4, column=0, columnspan=2, sticky="e", pady=(16, 0))
    ttk.Button(button_frame, text="Cancel", command=cancel).grid(row=0, column=0, padx=4)
    ttk.Button(button_frame, text="Save", command=submit).grid(row=0, column=1, padx=4)

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.bind("<Return>", lambda _event: submit())
    root.bind("<Escape>", lambda _event: cancel())
    root.update_idletasks()
    x = max(0, (root.winfo_screenwidth() - root.winfo_width()) // 2)
    y = max(0, (root.winfo_screenheight() - root.winfo_height()) // 2)
    root.geometry(f"+{x}+{y}")
    (password_entry if prefilled_account else account_entry).focus_set()
    root.mainloop()

    if result is None:
        raise CredentialsCancelledError("Gmail setup was cancelled")
    if not result.account or not result.app_password:
        raise CredentialsError("Gmail credentials were not provided")
    return result
