#!/usr/bin/env python3
"""
Graph App Role Manager
Cross-platform GUI for Windows and macOS.

Features:
- Authenticates with Microsoft Entra ID using MSAL device code flow.
- Searches Managed Identities / service principals by display name.
- Loads Microsoft Graph application permissions (app roles).
- Displays existing Microsoft Graph application permissions.
- Assigns selected app roles while avoiding duplicates.
- Removes selected app roles after explicit confirmation.

Required delegated permissions on the public-client App Registration:
- Application.Read.All
- AppRoleAssignment.ReadWrite.All

The signed-in administrator must also hold an appropriate Entra role.
"""

from __future__ import annotations

import json
import queue
import threading
import tkinter as tk
import webbrowser
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Any
from urllib.parse import quote

import msal
import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
SCOPES = [
    "Application.Read.All",
    "AppRoleAssignment.ReadWrite.All",
]


@dataclass(frozen=True)
class AppRole:
    id: str
    value: str
    display_name: str
    description: str


class GraphClient:
    def __init__(self) -> None:
        self.access_token: str | None = None
        self.graph_sp: dict[str, Any] | None = None

    @property
    def connected(self) -> bool:
        return bool(self.access_token)

    def _headers(self) -> dict[str, str]:
        if not self.access_token:
            raise RuntimeError("Not connected to Microsoft Graph.")
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

    def authenticate_device_code(
        self,
        tenant_id: str,
        client_id: str,
        callback,
    ) -> dict[str, Any]:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.PublicClientApplication(client_id=client_id, authority=authority)

        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(
                "Unable to start device-code authentication:\n"
                + json.dumps(flow, indent=2)
            )

        callback(flow)
        result = app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError(
                result.get("error_description")
                or result.get("error")
                or "Authentication failed."
            )

        self.access_token = result["access_token"]
        return result

    def get(self, path_or_url: str, params: dict[str, str] | None = None) -> dict[str, Any]:
        url = path_or_url if path_or_url.startswith("http") else f"{GRAPH_BASE}{path_or_url}"
        response = requests.get(url, headers=self._headers(), params=params, timeout=60)
        self._raise_for_graph_error(response)
        return response.json()

    def post(self, path: str, payload: dict[str, Any]) -> dict[str, Any]:
        response = requests.post(
            f"{GRAPH_BASE}{path}",
            headers=self._headers(),
            json=payload,
            timeout=60,
        )
        self._raise_for_graph_error(response)
        return response.json() if response.content else {}

    def delete(self, path: str) -> None:
        response = requests.delete(
            f"{GRAPH_BASE}{path}",
            headers=self._headers(),
            timeout=60,
        )
        self._raise_for_graph_error(response)

    @staticmethod
    def _raise_for_graph_error(response: requests.Response) -> None:
        if response.ok:
            return
        try:
            detail = response.json()
            message = detail.get("error", {}).get("message", response.text)
        except ValueError:
            message = response.text
        raise RuntimeError(f"Microsoft Graph HTTP {response.status_code}: {message}")

    def search_service_principals(self, display_name_prefix: str) -> list[dict[str, Any]]:
        escaped = display_name_prefix.replace("'", "''")
        result = self.get(
            "/servicePrincipals",
            params={
                "$filter": f"startsWith(displayName,'{escaped}')",
                "$select": "id,appId,displayName,servicePrincipalType,accountEnabled",
                "$top": "100",
            },
        )
        return sorted(result.get("value", []), key=lambda item: item.get("displayName", ""))

    def load_graph_app_roles(self) -> list[AppRole]:
        result = self.get(
            "/servicePrincipals",
            params={
                "$filter": f"appId eq '{GRAPH_APP_ID}'",
                "$select": "id,appId,displayName,appRoles",
            },
        )
        values = result.get("value", [])
        if not values:
            raise RuntimeError("Microsoft Graph service principal was not found.")

        self.graph_sp = values[0]
        roles: list[AppRole] = []
        for role in self.graph_sp.get("appRoles", []):
            if (
                role.get("isEnabled")
                and "Application" in role.get("allowedMemberTypes", [])
                and role.get("value")
            ):
                roles.append(
                    AppRole(
                        id=role["id"],
                        value=role["value"],
                        display_name=role.get("displayName", ""),
                        description=role.get("description", ""),
                    )
                )
        return sorted(roles, key=lambda role: role.value.lower())

    def get_graph_assignments(self, target_sp_id: str) -> list[dict[str, Any]]:
        if not self.graph_sp:
            self.load_graph_app_roles()

        items: list[dict[str, Any]] = []
        url: str | None = f"{GRAPH_BASE}/servicePrincipals/{target_sp_id}/appRoleAssignments?$top=999"

        while url:
            result = self.get(url)
            items.extend(result.get("value", []))
            url = result.get("@odata.nextLink")

        graph_resource_id = self.graph_sp["id"]
        return [item for item in items if item.get("resourceId") == graph_resource_id]

    def assign_app_role(self, target_sp_id: str, app_role_id: str) -> dict[str, Any]:
        if not self.graph_sp:
            raise RuntimeError("Microsoft Graph application roles are not loaded.")

        payload = {
            "principalId": target_sp_id,
            "resourceId": self.graph_sp["id"],
            "appRoleId": app_role_id,
        }
        return self.post(
            f"/servicePrincipals/{target_sp_id}/appRoleAssignments",
            payload,
        )

    def remove_app_role_assignment(
        self,
        target_sp_id: str,
        assignment_id: str,
    ) -> None:
        self.delete(
            f"/servicePrincipals/{target_sp_id}/appRoleAssignments/{assignment_id}"
        )


class GraphAppRoleManager(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Graph App Role Manager")
        self.geometry("1180x800")
        self.minsize(1000, 700)

        self.graph = GraphClient()
        self.roles: list[AppRole] = []
        self.filtered_roles: list[AppRole] = []
        self.selected_sp: dict[str, Any] | None = None
        self.current_assignments: dict[str, dict[str, Any]] = {}
        self.ui_queue: queue.Queue[tuple[str, Any]] = queue.Queue()

        self._configure_style()
        self._build_ui()
        self.after(100, self._process_ui_queue)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        available = style.theme_names()
        if "clam" in available:
            style.theme_use("clam")

        style.configure("Header.TFrame", background="#1f2937")
        style.configure(
            "HeaderTitle.TLabel",
            background="#1f2937",
            foreground="white",
            font=("Segoe UI", 20, "bold"),
        )
        style.configure(
            "HeaderSubtitle.TLabel",
            background="#1f2937",
            foreground="#d1d5db",
            font=("Segoe UI", 10),
        )
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Success.TButton", font=("Segoe UI", 10, "bold"))

    def _build_ui(self) -> None:
        header = ttk.Frame(self, style="Header.TFrame", padding=(20, 14))
        header.pack(fill="x")

        left = ttk.Frame(header, style="Header.TFrame")
        left.pack(side="left", fill="x", expand=True)
        ttk.Label(left, text="Graph App Role Manager", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(
            left,
            text="Assign Microsoft Graph application permissions to a Managed Identity",
            style="HeaderSubtitle.TLabel",
        ).pack(anchor="w")

        self.connect_button = ttk.Button(
            header,
            text="Connect to Microsoft Graph",
            style="Primary.TButton",
            command=self.connect,
        )
        self.connect_button.pack(side="right", padx=(15, 0))

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=14, pady=14)

        main_tab = ttk.Frame(notebook, padding=12)
        log_tab = ttk.Frame(notebook, padding=12)
        notebook.add(main_tab, text="Assign permissions")
        notebook.add(log_tab, text="Activity log")

        self._build_main_tab(main_tab)

        self.log_text = tk.Text(
            log_tab,
            wrap="none",
            background="#111827",
            foreground="#e5e7eb",
            insertbackground="white",
            font=("Menlo" if self.tk.call("tk", "windowingsystem") == "aqua" else "Consolas", 10),
        )
        self.log_text.pack(fill="both", expand=True)
        self.log("Graph App Role Manager started.")

    def _build_main_tab(self, parent: ttk.Frame) -> None:
        connection = ttk.LabelFrame(parent, text="Connection settings", padding=12)
        connection.pack(fill="x", pady=(0, 10))

        ttk.Label(connection, text="Tenant ID or tenant domain").grid(row=0, column=0, sticky="w")
        self.tenant_var = tk.StringVar()
        ttk.Entry(connection, textvariable=self.tenant_var, width=42).grid(
            row=1, column=0, sticky="ew", padx=(0, 12)
        )

        ttk.Label(connection, text="Public-client Application (Client) ID").grid(
            row=0, column=1, sticky="w"
        )
        self.client_var = tk.StringVar()
        ttk.Entry(connection, textvariable=self.client_var, width=42).grid(
            row=1, column=1, sticky="ew"
        )
        connection.columnconfigure(0, weight=1)
        connection.columnconfigure(1, weight=1)

        identity = ttk.LabelFrame(parent, text="1. Select the target Managed Identity", padding=12)
        identity.pack(fill="x", pady=(0, 10))

        row = ttk.Frame(identity)
        row.pack(fill="x")
        ttk.Label(row, text="Display name").pack(side="left")
        self.identity_name_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.identity_name_var).pack(
            side="left", fill="x", expand=True, padx=10
        )
        ttk.Button(row, text="Search", command=self.search_identity).pack(side="left")

        self.identity_combo = ttk.Combobox(identity, state="readonly")
        self.identity_combo.pack(fill="x", pady=(10, 8))
        self.identity_combo.bind("<<ComboboxSelected>>", self.identity_selected)

        summary = ttk.Frame(identity)
        summary.pack(fill="x")
        self.identity_summary_var = tk.StringVar(value="No identity selected.")
        ttk.Label(summary, textvariable=self.identity_summary_var).pack(anchor="w")

        body = ttk.Panedwindow(parent, orient=tk.HORIZONTAL)
        body.pack(fill="both", expand=True)

        permissions = ttk.LabelFrame(
            body, text="2. Select Microsoft Graph application permissions", padding=12
        )
        current = ttk.LabelFrame(
            body, text="3. Current Microsoft Graph assignments", padding=12
        )
        body.add(permissions, weight=3)
        body.add(current, weight=2)

        filter_row = ttk.Frame(permissions)
        filter_row.pack(fill="x")
        self.permission_filter_var = tk.StringVar()
        permission_entry = ttk.Entry(
            filter_row,
            textvariable=self.permission_filter_var,
        )
        permission_entry.pack(side="left", fill="x", expand=True)
        permission_entry.bind("<KeyRelease>", lambda _event: self.apply_permission_filter())
        ttk.Button(
            filter_row,
            text="Load permissions",
            command=self.load_permissions,
        ).pack(side="left", padx=(8, 0))

        list_frame = ttk.Frame(permissions)
        list_frame.pack(fill="both", expand=True, pady=10)

        self.permission_list = tk.Listbox(
            list_frame,
            selectmode=tk.MULTIPLE,
            activestyle="none",
            exportselection=False,
        )
        scrollbar = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=self.permission_list.yview,
        )
        self.permission_list.configure(yscrollcommand=scrollbar.set)
        self.permission_list.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        bottom = ttk.Frame(permissions)
        bottom.pack(fill="x")
        self.permission_count_var = tk.StringVar(value="0 permission(s) displayed")
        ttk.Label(bottom, textvariable=self.permission_count_var).pack(side="left")
        ttk.Button(
            bottom,
            text="Assign selected permissions",
            style="Success.TButton",
            command=self.assign_permissions,
        ).pack(side="right")

        refresh_row = ttk.Frame(current)
        refresh_row.pack(fill="x")
        ttk.Button(refresh_row, text="Refresh", command=self.refresh_assignments).pack(
            side="left"
        )
        ttk.Button(
            refresh_row,
            text="Remove selected",
            command=self.remove_permissions,
        ).pack(side="right")
        self.assigned_count_var = tk.StringVar(value="0 permission(s) assigned")
        ttk.Label(refresh_row, textvariable=self.assigned_count_var).pack(
            side="left", padx=10
        )

        columns = ("permission", "display")
        self.assignment_tree = ttk.Treeview(
            current,
            columns=columns,
            show="headings",
        )
        self.assignment_tree.heading("permission", text="Permission")
        self.assignment_tree.heading("display", text="Display name")
        self.assignment_tree.column("permission", width=260)
        self.assignment_tree.column("display", width=220)
        self.assignment_tree.pack(fill="both", expand=True, pady=(10, 0))

    def run_async(self, operation, *args) -> None:
        def worker() -> None:
            try:
                result = operation(*args)
                self.ui_queue.put(("success", result))
            except Exception as exc:  # noqa: BLE001
                self.ui_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _process_ui_queue(self) -> None:
        try:
            while True:
                event, payload = self.ui_queue.get_nowait()
                if event == "device_flow":
                    self.show_device_flow(payload)
                elif event == "connected":
                    self.on_connected(payload)
                elif event == "identities":
                    self.on_identities(payload)
                elif event == "roles":
                    self.on_roles(payload)
                elif event == "assignments":
                    self.on_assignments(payload)
                elif event == "assignment_complete":
                    self.on_assignment_complete(payload)
                elif event == "removal_complete":
                    self.on_removal_complete(payload)
                elif event == "error":
                    self.set_busy(False)
                    self.log(payload, "ERROR")
                    messagebox.showerror("Graph App Role Manager", payload)
        except queue.Empty:
            pass
        self.after(100, self._process_ui_queue)

    def set_busy(self, busy: bool) -> None:
        self.configure(cursor="watch" if busy else "")
        self.connect_button.configure(state="disabled" if busy else "normal")
        self.update_idletasks()

    def log(self, message: str, level: str = "INFO") -> None:
        from datetime import datetime

        line = f"[{datetime.now():%H:%M:%S}][{level}] {message}\n"
        self.log_text.insert("end", line)
        self.log_text.see("end")

    def connect(self) -> None:
        tenant_id = self.tenant_var.get().strip()
        client_id = self.client_var.get().strip()
        if not tenant_id or not client_id:
            messagebox.showwarning(
                "Missing configuration",
                "Enter the tenant ID/domain and the public-client Application ID.",
            )
            return

        self.set_busy(True)
        self.log("Starting device-code authentication...")

        def worker() -> None:
            try:
                result = self.graph.authenticate_device_code(
                    tenant_id,
                    client_id,
                    lambda flow: self.ui_queue.put(("device_flow", flow)),
                )
                self.ui_queue.put(("connected", result))
            except Exception as exc:  # noqa: BLE001
                self.ui_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def show_device_flow(self, flow: dict[str, Any]) -> None:
        verification_uri = flow.get("verification_uri") or flow.get("verification_uri_complete")
        code = flow.get("user_code", "")
        message = flow.get("message", "")

        self.log(message)
        self.clipboard_clear()
        self.clipboard_append(code)

        messagebox.showinfo(
            "Microsoft Graph authentication",
            f"{message}\n\nThe code has been copied to the clipboard.",
        )
        if verification_uri:
            webbrowser.open(verification_uri)

    def on_connected(self, result: dict[str, Any]) -> None:
        self.set_busy(False)
        account = result.get("id_token_claims", {}).get("preferred_username", "Connected")
        self.connect_button.configure(text=f"Connected: {account}")
        self.log(f"Connected as {account}.", "SUCCESS")
        self.load_permissions()

    def search_identity(self) -> None:
        name = self.identity_name_var.get().strip()
        if not self.graph.connected:
            messagebox.showwarning("Not connected", "Connect to Microsoft Graph first.")
            return
        if not name:
            messagebox.showwarning("Missing name", "Enter a display name.")
            return

        self.set_busy(True)
        self.log(f"Searching service principals matching '{name}'...")

        def worker() -> None:
            try:
                values = self.graph.search_service_principals(name)
                self.ui_queue.put(("identities", values))
            except Exception as exc:  # noqa: BLE001
                self.ui_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def on_identities(self, values: list[dict[str, Any]]) -> None:
        self.set_busy(False)
        self.identity_results = values
        labels = [
            f"{item.get('displayName')} | {item.get('servicePrincipalType')} | {item.get('id')}"
            for item in values
        ]
        self.identity_combo["values"] = labels
        if values:
            self.identity_combo.current(0)
            self.identity_selected()
            self.log(f"Found {len(values)} matching service principal(s).", "SUCCESS")
        else:
            self.selected_sp = None
            self.identity_summary_var.set("No matching service principal found.")
            self.log("No matching service principal found.", "WARNING")

    def identity_selected(self, _event=None) -> None:
        index = self.identity_combo.current()
        if index < 0:
            return
        self.selected_sp = self.identity_results[index]
        self.identity_summary_var.set(
            f"Name: {self.selected_sp.get('displayName')}   |   "
            f"Object ID: {self.selected_sp.get('id')}   |   "
            f"App ID: {self.selected_sp.get('appId')}"
        )
        self.log(f"Selected target: {self.selected_sp.get('displayName')}")
        self.refresh_assignments()

    def load_permissions(self) -> None:
        if not self.graph.connected:
            return
        self.set_busy(True)
        self.log("Loading Microsoft Graph application permissions...")

        def worker() -> None:
            try:
                roles = self.graph.load_graph_app_roles()
                self.ui_queue.put(("roles", roles))
            except Exception as exc:  # noqa: BLE001
                self.ui_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def on_roles(self, roles: list[AppRole]) -> None:
        self.set_busy(False)
        self.roles = roles
        self.apply_permission_filter()
        self.log(f"Loaded {len(roles)} application permissions.", "SUCCESS")

    def apply_permission_filter(self) -> None:
        term = self.permission_filter_var.get().strip().lower()
        self.filtered_roles = [
            role
            for role in self.roles
            if not term
            or term in role.value.lower()
            or term in role.display_name.lower()
            or term in role.description.lower()
        ]

        self.permission_list.delete(0, "end")
        for role in self.filtered_roles:
            self.permission_list.insert(
                "end",
                f"{role.value} — {role.display_name}",
            )
        self.permission_count_var.set(
            f"{len(self.filtered_roles)} permission(s) displayed"
        )

    def refresh_assignments(self) -> None:
        if not self.selected_sp or not self.graph.connected:
            return
        target_sp_id = self.selected_sp["id"]
        self.set_busy(True)
        self.log("Loading current Microsoft Graph assignments...")

        def worker() -> None:
            try:
                assignments = self.graph.get_graph_assignments(target_sp_id)
                self.ui_queue.put(("assignments", assignments))
            except Exception as exc:  # noqa: BLE001
                self.ui_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def on_assignments(self, assignments: list[dict[str, Any]]) -> None:
        self.set_busy(False)
        self.current_assignments = {
            assignment["id"]: assignment
            for assignment in assignments
            if assignment.get("id")
        }
        for item in self.assignment_tree.get_children():
            self.assignment_tree.delete(item)

        role_by_id = {role.id: role for role in self.roles}
        for assignment in assignments:
            role = role_by_id.get(assignment.get("appRoleId"))
            permission = role.value if role else assignment.get("appRoleId", "Unknown")
            display = role.display_name if role else ""
            self.assignment_tree.insert(
                "",
                "end",
                iid=assignment.get("id"),
                values=(permission, display),
            )

        self.assigned_count_var.set(f"{len(assignments)} permission(s) assigned")
        self.log(f"Current assignments loaded: {len(assignments)}.", "SUCCESS")

    def assign_permissions(self) -> None:
        if not self.selected_sp:
            messagebox.showwarning("No target", "Select a Managed Identity first.")
            return

        target_sp = dict(self.selected_sp)
        indexes = self.permission_list.curselection()
        selected_roles = [self.filtered_roles[index] for index in indexes]
        if not selected_roles:
            messagebox.showwarning(
                "No permissions selected",
                "Select at least one Microsoft Graph application permission.",
            )
            return

        permission_names = "\n".join(role.value for role in selected_roles)
        confirmed = messagebox.askyesno(
            "Confirm permission assignment",
            f"Target:\n{target_sp.get('displayName')}\n\n"
            f"Permissions:\n{permission_names}\n\nProceed?",
        )
        if not confirmed:
            return

        self.set_busy(True)

        def worker() -> None:
            try:
                existing = self.graph.get_graph_assignments(target_sp["id"])
                existing_ids = {item.get("appRoleId") for item in existing}
                summary = {"assigned": [], "skipped": [], "failed": []}

                for role in selected_roles:
                    if role.id in existing_ids:
                        summary["skipped"].append(role.value)
                        continue
                    try:
                        self.graph.assign_app_role(target_sp["id"], role.id)
                        summary["assigned"].append(role.value)
                    except Exception as exc:  # noqa: BLE001
                        summary["failed"].append(f"{role.value}: {exc}")

                self.ui_queue.put(("assignment_complete", summary))
            except Exception as exc:  # noqa: BLE001
                self.ui_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def on_assignment_complete(self, summary: dict[str, list[str]]) -> None:
        self.set_busy(False)

        for permission in summary["assigned"]:
            self.log(f"Assigned successfully: {permission}", "SUCCESS")
        for permission in summary["skipped"]:
            self.log(f"Already assigned: {permission}", "WARNING")
        for failure in summary["failed"]:
            self.log(f"Failed: {failure}", "ERROR")

        messagebox.showinfo(
            "Graph App Role Manager",
            "Completed.\n\n"
            f"Assigned: {len(summary['assigned'])}\n"
            f"Already present: {len(summary['skipped'])}\n"
            f"Failed: {len(summary['failed'])}",
        )
        self.refresh_assignments()

    def remove_permissions(self) -> None:
        if not self.selected_sp:
            messagebox.showwarning("No target", "Select a Managed Identity first.")
            return

        target_sp = dict(self.selected_sp)
        selected_ids = list(self.assignment_tree.selection())
        selected_assignments = [
            self.current_assignments[assignment_id]
            for assignment_id in selected_ids
            if assignment_id in self.current_assignments
        ]
        if not selected_assignments:
            messagebox.showwarning(
                "No permissions selected",
                "Select at least one assigned Microsoft Graph permission.",
            )
            return

        role_by_id = {role.id: role for role in self.roles}
        permission_names = "\n".join(
            role_by_id[assignment["appRoleId"]].value
            if assignment.get("appRoleId") in role_by_id
            else assignment.get("appRoleId", "Unknown")
            for assignment in selected_assignments
        )
        confirmed = messagebox.askyesno(
            "Confirm permission removal",
            f"Target:\n{target_sp.get('displayName')}\n\n"
            f"Permissions to remove:\n{permission_names}\n\nProceed?",
        )
        if not confirmed:
            self.log("Permission removal cancelled by the user.", "WARNING")
            return

        self.set_busy(True)

        def worker() -> None:
            summary = {"removed": [], "failed": []}
            for assignment in selected_assignments:
                role = role_by_id.get(assignment.get("appRoleId"))
                permission = (
                    role.value
                    if role
                    else assignment.get("appRoleId", "Unknown")
                )
                try:
                    self.graph.remove_app_role_assignment(
                        target_sp["id"],
                        assignment["id"],
                    )
                    summary["removed"].append(permission)
                except Exception as exc:  # noqa: BLE001
                    summary["failed"].append(f"{permission}: {exc}")

            self.ui_queue.put(("removal_complete", summary))

        threading.Thread(target=worker, daemon=True).start()

    def on_removal_complete(self, summary: dict[str, list[str]]) -> None:
        self.set_busy(False)

        for permission in summary["removed"]:
            self.log(f"Removed successfully: {permission}", "SUCCESS")
        for failure in summary["failed"]:
            self.log(f"Failed: {failure}", "ERROR")

        messagebox.showinfo(
            "Graph App Role Manager",
            "Completed.\n\n"
            f"Removed: {len(summary['removed'])}\n"
            f"Failed: {len(summary['failed'])}",
        )
        self.refresh_assignments()


def main() -> None:
    app = GraphAppRoleManager()
    app.mainloop()


if __name__ == "__main__":
    main()
