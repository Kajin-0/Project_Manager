import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import csv

STATUS_OPTIONS = ["Not Started", "Completed", "In Progress", "Aborted", "Paused"]

COLORS = {
    'background': '#F0F4F8',   # Light blue-gray background
    'primary': '#2C3E50',      # Dark blue-gray for headers and accents
    'secondary': '#34495E',    # Slightly lighter blue-gray
    'highlight': '#3498DB',    # Bright blue for interactive elements
    'text': '#2C3E50',         # Dark text color
    'status_colors': {
        'Not Started': '#E74C3C',  # Red
        'Completed': '#2ECC71',    # Green
        'In Progress': '#F39C12',  # Orange
        'Aborted': '#95A5A6',      # Gray
        'Paused': '#9B59B6'        # Purple
    }
}

class ProjectManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Professional Project Manager")
        self.root.configure(bg=COLORS['background'])
        self.global_personnel = set()
        self.projects = []            # Active (non-completed) projects
        self.completed_projects = []  # Completed projects

        # Setup style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TFrame", background=COLORS['background'])
        style.configure("TLabel",
                        background=COLORS['background'],
                        foreground=COLORS['text'],
                        font=('Segoe UI', 10))
        style.configure("Title.TLabel",
                        font=('Segoe UI', 16, 'bold'),
                        foreground=COLORS['primary'])
        style.configure("TButton",
                        background=COLORS['highlight'],
                        foreground='white',
                        font=('Segoe UI', 10, 'bold'))
        style.map('TButton',
                  background=[('active', COLORS['secondary']),
                              ('pressed', COLORS['primary'])])

        # Menu
        menubar = tk.Menu(root, bg=COLORS['background'], fg=COLORS['text'], tearoff=0)
        filemenu = tk.Menu(menubar, tearoff=0, bg='white', fg=COLORS['text'])
        filemenu.add_command(label="Export CSV", command=self.export_csv)
        filemenu.add_command(label="Import CSV", command=self.import_csv)
        filemenu.add_separator()
        filemenu.add_command(label="Manage Global Personnel", command=self.manage_global_personnel)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=root.quit)
        menubar.add_cascade(label="File", menu=filemenu)
        root.config(menu=menubar)

        main_frame = ttk.Frame(root, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(main_frame, text="Project Management Dashboard", style="Title.TLabel")
        title_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0,20))

        # Notebook for Active and Completed Projects
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, columnspan=3, sticky=tk.NSEW)

        # Frames for Active and Completed tabs
        self.active_frame = ttk.Frame(self.notebook)
        self.completed_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.active_frame, text="Active Projects")
        self.notebook.add(self.completed_frame, text="Completed Projects")

        # ACTIVE PROJECTS TAB
        # Left side: Project list and details
        left_frame = ttk.Frame(self.active_frame)
        left_frame.grid(row=0, column=0, sticky=tk.NSEW)

        # Right side: Buttons
        right_frame = ttk.Frame(self.active_frame)
        right_frame.grid(row=0, column=2, sticky=tk.N, padx=10)

        # Middle frame for arrows
        middle_frame = ttk.Frame(self.active_frame)
        middle_frame.grid(row=0, column=1, sticky=tk.N)

        self.active_frame.columnconfigure(0, weight=1)
        self.active_frame.rowconfigure(0, weight=1)

        ttk.Label(left_frame, text="Projects:").pack(anchor=tk.W)
        self.project_listbox = tk.Listbox(
            left_frame,
            height=15,
            width=50,
            bg='white',
            fg=COLORS['text'],
            selectbackground=COLORS['highlight'],
            font=('Segoe UI', 10)
        )
        self.project_listbox.pack(pady=(5, 10), fill=tk.BOTH, expand=True)
        self.project_listbox.bind("<<ListboxSelect>>", self.on_project_select)

        details_frame = ttk.LabelFrame(left_frame, text="Selected Project Details", padding="10 10 10 10")
        details_frame.pack(fill=tk.X)
        self.details_label = ttk.Label(details_frame, text="No project selected.")
        self.details_label.pack(anchor=tk.W)

        # Up/Down arrow buttons with nicer look
        # Using unicode arrows ▲ and ▼
        self.move_up_button = ttk.Button(middle_frame, text="▲", command=self.move_project_up, width=3)
        self.move_up_button.grid(row=0, column=0, pady=5, padx=5, sticky=tk.EW)

        self.move_down_button = ttk.Button(middle_frame, text="▼", command=self.move_project_down, width=3)
        self.move_down_button.grid(row=1, column=0, pady=5, padx=5, sticky=tk.EW)

        button_frame = ttk.Frame(right_frame)
        button_frame.pack(pady=10)

        self.add_project_button = ttk.Button(button_frame, text="Add Project", command=self.add_project)
        self.add_project_button.grid(row=0, column=0, pady=5, padx=5, sticky=tk.EW)

        self.edit_project_button = ttk.Button(button_frame, text="Edit Project", command=self.edit_project)
        self.edit_project_button.grid(row=1, column=0, pady=5, padx=5, sticky=tk.EW)

        self.delete_project_button = ttk.Button(button_frame, text="Delete Project", command=self.delete_project)
        self.delete_project_button.grid(row=2, column=0, pady=5, padx=5, sticky=tk.EW)

        self.manage_subprocess_button = ttk.Button(button_frame, text="Manage Sub-Processes", command=self.manage_subprocesses)
        self.manage_subprocess_button.grid(row=3, column=0, pady=5, padx=5, sticky=tk.EW)

        self.manage_personnel_button = ttk.Button(button_frame, text="Manage Project Personnel", command=self.manage_project_personnel)
        self.manage_personnel_button.grid(row=4, column=0, pady=5, padx=5, sticky=tk.EW)

        self.view_details_button = ttk.Button(button_frame, text="View Full Details", command=self.view_details)
        self.view_details_button.grid(row=5, column=0, pady=5, padx=5, sticky=tk.EW)

        # COMPLETED PROJECTS TAB
        # Similar layout but for completed projects
        self.completed_frame.columnconfigure(0, weight=1)
        self.completed_frame.rowconfigure(0, weight=1)

        completed_main_frame = ttk.Frame(self.completed_frame)
        completed_main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(completed_main_frame, text="Completed Projects:").grid(row=0, column=0, columnspan=3, sticky=tk.W)

        self.completed_listbox = tk.Listbox(
            completed_main_frame,
            height=10,
            width=50,
            bg='white',
            fg=COLORS['text'],
            selectbackground=COLORS['highlight'],
            font=('Segoe UI', 10)
        )
        self.completed_listbox.grid(row=1, column=0, columnspan=3, pady=(5,10), sticky=tk.NSEW)

        # Revert status with a dropdown
        ttk.Label(completed_main_frame, text="Revert Status To:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        self.revert_status_var = tk.StringVar(value="Not Started")
        self.revert_status_combobox = ttk.Combobox(completed_main_frame, textvariable=self.revert_status_var, values=[s for s in STATUS_OPTIONS if s != "Completed"], state='readonly')
        self.revert_status_combobox.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        self.revert_status_combobox.set("Not Started")

        revert_button = ttk.Button(completed_main_frame, text="Revert", command=self.revert_completed_project)
        revert_button.grid(row=2, column=2, padx=5, pady=5, sticky=tk.EW)

        self.completed_frame.rowconfigure(1, weight=1)

        # Configure main_frame's weight
        main_frame.rowconfigure(1, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(2, weight=0)

        self.update_project_list()
        self.update_completed_list()

    def on_project_select(self, event):
        selected_index = self.get_selected_project_index(quiet=True)
        if selected_index is None:
            self.details_label.config(text="No project selected.")
            return
        project = self.projects[selected_index]
        details_text = (
            f"Name: {project['name']}\n"
            f"Status: {project['status']}\n"
            f"Start Date: {project['start_date']}\n"
            f"End Date: {project['end_date']}"
        )
        self.details_label.config(text=details_text)

    def add_project(self):
        dialog = ProjectDialog(self.root, "Add Project", allow_name_edit=True, allow_date_edit=True)
        self.root.wait_window(dialog.top)
        if dialog.result is not None:
            name, status, start_date, end_date = dialog.result
            project = {
                "name": name,
                "status": status,
                "start_date": start_date,
                "end_date": end_date,
                "personnel": [],
                "subprocesses": []
            }
            if status == "Completed":
                self.completed_projects.append(project)
                self.update_completed_list()
            else:
                self.projects.append(project)
                self.update_project_list()

    def edit_project(self):
        selected_index = self.get_selected_project_index()
        if selected_index is None:
            return
        project = self.projects[selected_index]
        dialog = ProjectDialog(
            self.root,
            "Edit Project",
            initial_name=project["name"],
            initial_status=project["status"],
            initial_start_date=project["start_date"],
            initial_end_date=project["end_date"],
            allow_name_edit=False,
            allow_date_edit=True
        )
        self.root.wait_window(dialog.top)
        if dialog.result is not None:
            _, new_status, new_start, new_end = dialog.result
            project["status"] = new_status
            project["start_date"] = new_start
            project["end_date"] = new_end
            # If now completed, move it to completed_projects
            if new_status == "Completed":
                self.completed_projects.append(project)
                self.projects.pop(selected_index)
                self.update_completed_list()
            self.update_project_list()

    def delete_project(self):
        selected_index = self.get_selected_project_index()
        if selected_index is None:
            return
        project = self.projects[selected_index]
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the project '{project['name']}'?", parent=self.root):
            self.projects.pop(selected_index)
            self.update_project_list()

    def manage_subprocesses(self):
        selected_index = self.get_selected_project_index()
        if selected_index is None:
            return
        project = self.projects[selected_index]
        dialog = ManageSubProcessesDialog(self.root, project["subprocesses"], self.global_personnel)
        self.root.wait_window(dialog.top)
        self.projects[selected_index]["subprocesses"] = dialog.subprocesses

    def manage_project_personnel(self):
        selected_index = self.get_selected_project_index()
        if selected_index is None:
            return
        project = self.projects[selected_index]
        dialog = ManageAssignmentsDialog(self.root, project["personnel"], self.global_personnel, "Project Personnel")
        self.root.wait_window(dialog.top)
        project["personnel"] = dialog.assignments

    def view_details(self):
        selected_index = self.get_selected_project_index()
        if selected_index is None:
            return
        project = self.projects[selected_index]
        self.show_project_details_window(project)

    def show_project_details_window(self, project):
        # Create a new top-level window for project details
        details_window = tk.Toplevel(self.root)
        details_window.title("Project Details")
        details_window.configure(bg=COLORS['background'])
        details_window.transient(self.root)
        details_window.grab_set()

        # Style a frame inside the top-level for padding
        main_frame = ttk.Frame(details_window, padding="20 20 20 20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text=f"Project: {project['name']}", style="Title.TLabel")
        title_label.pack(anchor=tk.W, pady=(0,10))

        # Notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        # ---------------- Overview Tab ----------------
        overview_frame = ttk.Frame(notebook)
        notebook.add(overview_frame, text="Overview")

        # Grid layout for overview
        ttk.Label(overview_frame, text="Status:", font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Label(overview_frame, text=project['status']).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(overview_frame, text="Start Date:", font=('Segoe UI', 10, 'bold')).grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Label(overview_frame, text=project['start_date']).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Label(overview_frame, text="End Date:", font=('Segoe UI', 10, 'bold')).grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Label(overview_frame, text=project['end_date']).grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)

        # ---------------- Sub-Processes Tab ----------------
        subprocess_frame = ttk.Frame(notebook)
        notebook.add(subprocess_frame, text="Sub-Processes")

        sp_columns = ("Name", "Status", "Start", "End")
        sp_tree = ttk.Treeview(subprocess_frame, columns=sp_columns, show='headings', height=10)
        for col in sp_columns:
            sp_tree.heading(col, text=col)
            # Assign some reasonable widths
            if col == "Name":
                sp_tree.column(col, width=150)
            else:
                sp_tree.column(col, width=100)

        sp_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Populate sub-process tree
        for sp in project["subprocesses"]:
            sp_id = sp_tree.insert("", tk.END, values=(sp['name'], sp['status'], sp['start_date'], sp['end_date']))
            # If the sub-process has personnel, add them as child items
            if sp["personnel"]:
                for (pname, role) in sp["personnel"]:
                    sp_tree.insert(sp_id, tk.END, values=(f"Personnel: {pname}", role, "", ""))

        # ---------------- Personnel Tab ----------------
        # Just project-level personnel here
        personnel_frame = ttk.Frame(notebook)
        notebook.add(personnel_frame, text="Personnel")

        p_columns = ("Name", "Role")
        p_tree = ttk.Treeview(personnel_frame, columns=p_columns, show='headings', height=10)
        for col in p_columns:
            p_tree.heading(col, text=col)
            p_tree.column(col, width=150)

        p_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Populate personnel tree with project-level personnel
        for (pname, role) in project["personnel"]:
            p_tree.insert("", tk.END, values=(pname, role))

        # Make the details window slightly larger
        details_window.geometry("600x400")

    def revert_completed_project(self):
        sel = self.completed_listbox.curselection()
        if not sel:
            messagebox.showerror("Error", "No completed project selected.", parent=self.root)
            return
        idx = sel[0]
        project = self.completed_projects[idx]
        new_status = self.revert_status_var.get()
        # Validate
        if new_status not in STATUS_OPTIONS or new_status == "Completed":
            messagebox.showerror("Error", "Invalid status selected.", parent=self.root)
            return

        project["status"] = new_status
        # Move project back to active
        self.projects.append(project)
        self.completed_projects.pop(idx)
        self.update_project_list()
        self.update_completed_list()

    def update_project_list(self):
        self.project_listbox.delete(0, tk.END)
        for project in self.projects:
            self.project_listbox.insert(tk.END, f"{project['name']} ({project['status']})")
        self.on_project_select(None)

    def update_completed_list(self):
        self.completed_listbox.delete(0, tk.END)
        for p in self.completed_projects:
            self.completed_listbox.insert(tk.END, f"{p['name']} ({p['status']})")

    def get_selected_project_index(self, quiet=False):
        selected_indices = self.project_listbox.curselection()
        if not selected_indices:
            if not quiet:
                messagebox.showerror("Error", "No project selected.")
            return None
        return selected_indices[0]

    def manage_global_personnel(self):
        dialog = GlobalPersonnelDialog(self.root, self.global_personnel)
        self.root.wait_window(dialog.top)
        self.cleanup_deleted_personnel()

    def cleanup_deleted_personnel(self):
        def filter_assignments(assignments):
            return [(p, r) for (p, r) in assignments if p in self.global_personnel]

        for project in self.projects:
            project["personnel"] = filter_assignments(project["personnel"])
            for sp in project["subprocesses"]:
                sp["personnel"] = filter_assignments(sp["personnel"])

        for cproject in self.completed_projects:
            cproject["personnel"] = filter_assignments(cproject["personnel"])
            for sp in cproject["subprocesses"]:
                sp["personnel"] = filter_assignments(sp["personnel"])

    def export_csv(self):
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if not filename:
            return

        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Global personnel
            for person_name in sorted(self.global_personnel):
                writer.writerow(["PERSON", person_name])

            # Active projects
            for project in self.projects:
                writer.writerow(["PROJECT", project["name"], project["status"], project["start_date"], project["end_date"]])
                for (pname, role) in project["personnel"]:
                    writer.writerow(["PPERSONNEL", project["name"], pname, role])
                for sp in project["subprocesses"]:
                    writer.writerow(["SUBPROCESS", project["name"], sp["name"], sp["status"], sp["start_date"], sp["end_date"]])
                    for (pname, role) in sp["personnel"]:
                        writer.writerow(["SPERSONNEL", project["name"], sp["name"], pname, role])

            # Completed projects
            for cproject in self.completed_projects:
                writer.writerow(["PROJECT", cproject["name"], cproject["status"], cproject["start_date"], cproject["end_date"]])
                for (pname, role) in cproject["personnel"]:
                    writer.writerow(["PPERSONNEL", cproject["name"], pname, role])
                for sp in cproject["subprocesses"]:
                    writer.writerow(["SUBPROCESS", cproject["name"], sp["name"], sp["status"], sp["start_date"], sp["end_date"]])
                    for (pname, role) in sp["personnel"]:
                        writer.writerow(["SPERSONNEL", cproject["name"], sp["name"], pname, role])

        messagebox.showinfo("Success", f"Projects exported to {filename}")

    def import_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not filename:
            return

        self.projects.clear()
        self.completed_projects.clear()
        self.global_personnel.clear()

        projects_dict = {}

        with open(filename, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if not row:
                    continue
                rtype = row[0]
                if rtype == "PERSON":
                    _, pname = row
                    self.global_personnel.add(pname)
                elif rtype == "PROJECT":
                    _, pname, status, sdate, edate = row
                    projects_dict[pname] = {
                        "name": pname,
                        "status": status,
                        "start_date": sdate,
                        "end_date": edate,
                        "personnel": [],
                        "subprocesses": []
                    }
                elif rtype == "PPERSONNEL":
                    _, pname, pername, perrole = row
                    if pername not in self.global_personnel:
                        self.global_personnel.add(pername)
                    if pname in projects_dict:
                        projects_dict[pname]["personnel"].append((pername, perrole))
                elif rtype == "SUBPROCESS":
                    _, pname, spname, spstatus, spsdate, spedate = row
                    if pname in projects_dict:
                        projects_dict[pname]["subprocesses"].append({
                            "name": spname,
                            "status": spstatus,
                            "start_date": spsdate,
                            "end_date": spedate,
                            "personnel": []
                        })
                elif rtype == "SPERSONNEL":
                    _, pname, spname, pername, perrole = row
                    if pername not in self.global_personnel:
                        self.global_personnel.add(pername)
                    if pname in projects_dict:
                        for sp in projects_dict[pname]["subprocesses"]:
                            if sp["name"] == spname:
                                sp["personnel"].append((pername, perrole))
                                break

        # Separate completed from active
        for p in projects_dict.values():
            if p["status"] == "Completed":
                self.completed_projects.append(p)
            else:
                self.projects.append(p)

        self.update_project_list()
        self.update_completed_list()
        messagebox.showinfo("Success", f"Projects imported from {filename}")

    def move_project_up(self):
        idx = self.get_selected_project_index()
        if idx is None:
            return
        if idx > 0:
            self.projects[idx], self.projects[idx-1] = self.projects[idx-1], self.projects[idx]
            self.update_project_list()
            self.project_listbox.selection_set(idx-1)

    def move_project_down(self):
        idx = self.get_selected_project_index()
        if idx is None:
            return
        if idx < len(self.projects)-1:
            self.projects[idx], self.projects[idx+1] = self.projects[idx+1], self.projects[idx]
            self.update_project_list()
            self.project_listbox.selection_set(idx+1)


class ProjectDialog:
    def __init__(self, parent, title, initial_name="", initial_status="Not Started",
                 initial_start_date="", initial_end_date="",
                 allow_name_edit=True, allow_date_edit=True):
        self.result = None

        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.configure(bg=COLORS['background'])
        self.top.transient(parent)
        self.top.grab_set()

        frame = ttk.Frame(self.top, padding="10 10 10 10")
        frame.grid(row=0, column=0, sticky=tk.NSEW)

        ttk.Label(frame, text="Project Name:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.name_entry = ttk.Entry(frame)
        self.name_entry.grid(row=0, column=1, padx=10, pady=5)
        self.name_entry.insert(0, initial_name)
        if not allow_name_edit:
            self.name_entry.config(state='disabled')

        ttk.Label(frame, text="Status:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.status_var = tk.StringVar(value=initial_status)
        self.status_combobox = ttk.Combobox(frame, textvariable=self.status_var, values=STATUS_OPTIONS, state='readonly')
        self.status_combobox.grid(row=1, column=1, padx=10, pady=5)
        self.status_combobox.set(initial_status)

        ttk.Label(frame, text="Start Date:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
        self.start_date_entry = ttk.Entry(frame)
        self.start_date_entry.grid(row=2, column=1, padx=10, pady=5)
        self.start_date_entry.insert(0, initial_start_date)
        if not allow_date_edit:
            self.start_date_entry.config(state='disabled')

        ttk.Label(frame, text="End Date:").grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
        self.end_date_entry = ttk.Entry(frame)
        self.end_date_entry.grid(row=3, column=1, padx=10, pady=5)
        self.end_date_entry.insert(0, initial_end_date)
        if not allow_date_edit:
            self.end_date_entry.config(state='disabled')

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="OK", command=self.on_ok).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.top.destroy).grid(row=0, column=1, padx=5)

    def on_ok(self):
        name = self.name_entry.get().strip()
        status = self.status_var.get()
        start_date = self.start_date_entry.get().strip()
        end_date = self.end_date_entry.get().strip()

        if not name:
            messagebox.showerror("Error", "Name cannot be empty.", parent=self.top)
            return
        if status not in STATUS_OPTIONS:
            messagebox.showerror("Error", "Invalid status.", parent=self.top)
            return

        self.result = (name, status, start_date, end_date)
        self.top.destroy()


class ManagePersonnelDialog:
    def __init__(self, parent, personnel_set):
        self.top = tk.Toplevel(parent)
        self.top.title("Global Personnel")
        self.top.configure(bg=COLORS['background'])
        self.top.transient(parent)
        self.top.grab_set()

        self.personnel_set = personnel_set

        frame = ttk.Frame(self.top, padding="10 10 10 10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Global Personnel:").grid(row=0, column=0, columnspan=3, sticky=tk.W)
        self.personnel_listbox = tk.Listbox(frame, height=10, width=40, bg='white', fg=COLORS['text'],
                                            selectbackground=COLORS['highlight'], font=('Segoe UI', 10))
        self.personnel_listbox.grid(row=1, column=0, columnspan=3, pady=(5,10))
        self.update_personnel_list()

        add_button = ttk.Button(frame, text="Add", command=self.add_personnel)
        add_button.grid(row=2, column=0, sticky=tk.EW, padx=5, pady=5)

        edit_button = ttk.Button(frame, text="Edit", command=self.edit_personnel)
        edit_button.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)

        remove_button = ttk.Button(frame, text="Remove", command=self.remove_personnel)
        remove_button.grid(row=2, column=2, sticky=tk.EW, padx=5, pady=5)

        close_button = ttk.Button(frame, text="Close", command=self.top.destroy)
        close_button.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)

    def update_personnel_list(self):
        self.personnel_listbox.delete(0, tk.END)
        for p in sorted(self.personnel_set):
            self.personnel_listbox.insert(tk.END, p)

    def get_selected_personnel(self):
        sel = self.personnel_listbox.curselection()
        if not sel:
            messagebox.showerror("Error", "No personnel selected.", parent=self.top)
            return None
        return self.personnel_listbox.get(sel[0])

    def add_personnel(self):
        name = simpledialog.askstring("Personnel Name", "Enter personnel name:", parent=self.top)
        if not name:
            return
        if name in self.personnel_set:
            messagebox.showerror("Error", "A person with this name already exists.", parent=self.top)
            return
        self.personnel_set.add(name)
        self.update_personnel_list()

    def edit_personnel(self):
        old_name = self.get_selected_personnel()
        if old_name is None:
            return
        new_name = simpledialog.askstring("Edit Personnel Name", f"Current: {old_name}\nEnter new name:", parent=self.top, initialvalue=old_name)
        if not new_name:
            return
        if new_name != old_name and new_name in self.personnel_set:
            messagebox.showerror("Error", "A person with this name already exists.", parent=self.top)
            return
        if new_name == old_name:
            return

        self.personnel_set.remove(old_name)
        self.personnel_set.add(new_name)
        # Rename globally
        app = self.get_app()
        for project in app.projects:
            project["personnel"] = [(new_name if p == old_name else p, r) for (p, r) in project["personnel"]]
            for sp in project["subprocesses"]:
                sp["personnel"] = [(new_name if p == old_name else p, r) for (p, r) in sp["personnel"]]

        for cproject in app.completed_projects:
            cproject["personnel"] = [(new_name if p == old_name else p, r) for (p, r) in cproject["personnel"]]
            for sp in cproject["subprocesses"]:
                sp["personnel"] = [(new_name if p == old_name else p, r) for (p, r) in sp["personnel"]]

        self.update_personnel_list()

    def remove_personnel(self):
        name = self.get_selected_personnel()
        if name is None:
            return
        if messagebox.askyesno("Confirm", f"Are you sure you want to delete {name}? They will be removed from all assignments.", parent=self.top):
            self.personnel_set.remove(name)
            app = self.get_app()
            app.cleanup_deleted_personnel()
            self.update_personnel_list()

    def get_app(self):
        app = self.top.master.nametowidget(self.top.master.winfo_parent())
        while not isinstance(app, ProjectManagerApp):
            app = app.master
        return app


class ManageAssignmentsDialog:
    def __init__(self, parent, assignments, global_personnel, title="Manage Assignments"):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.configure(bg=COLORS['background'])
        self.top.transient(parent)
        self.top.grab_set()

        self.assignments = assignments
        self.global_personnel = global_personnel

        frame = ttk.Frame(self.top, padding="10 10 10 10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text=title + ":").grid(row=0, column=0, columnspan=3, sticky=tk.W)
        self.listbox = tk.Listbox(frame, height=10, width=50, bg='white', fg=COLORS['text'],
                                  selectbackground=COLORS['highlight'], font=('Segoe UI', 10))
        self.listbox.grid(row=1, column=0, columnspan=3, pady=(5,10))
        self.update_list()

        add_button = ttk.Button(frame, text="Add", command=self.add_assignment)
        add_button.grid(row=2, column=0, sticky=tk.EW, padx=5, pady=5)

        edit_button = ttk.Button(frame, text="Edit", command=self.edit_assignment)
        edit_button.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)

        remove_button = ttk.Button(frame, text="Remove", command=self.remove_assignment)
        remove_button.grid(row=2, column=2, sticky=tk.EW, padx=5, pady=5)

        close_button = ttk.Button(frame, text="Close", command=self.top.destroy)
        close_button.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)

    def update_list(self):
        self.listbox.delete(0, tk.END)
        for (p, r) in self.assignments:
            self.listbox.insert(tk.END, f"{p} ({r})")

    def get_selected_index(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showerror("Error", "No assignment selected.", parent=self.top)
            return None
        return sel[0]

    def add_assignment(self):
        if not self.global_personnel:
            messagebox.showerror("Error", "No global personnel available. Add global personnel first.", parent=self.top)
            return
        dialog = AssignmentDialog(self.top, self.global_personnel)
        self.top.wait_window(dialog.top)
        if dialog.result is not None:
            pname, role = dialog.result
            self.assignments.append((pname, role))
            self.update_list()

    def edit_assignment(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        (pname, role) = self.assignments[idx]
        if not self.global_personnel:
            messagebox.showerror("Error", "No global personnel available.", parent=self.top)
            return
        dialog = AssignmentDialog(self.top, self.global_personnel, initial_person=pname, initial_role=role)
        self.top.wait_window(dialog.top)
        if dialog.result is not None:
            new_pname, new_role = dialog.result
            self.assignments[idx] = (new_pname, new_role)
            self.update_list()

    def remove_assignment(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        self.assignments.pop(idx)
        self.update_list()


class AssignmentDialog:
    def __init__(self, parent, global_personnel, initial_person=None, initial_role=""):
        self.result = None
        self.top = tk.Toplevel(parent)
        self.top.title("Assign Personnel")
        self.top.configure(bg=COLORS['background'])
        self.top.transient(parent)
        self.top.grab_set()

        global_personnel_list = sorted(global_personnel)

        frame = ttk.Frame(self.top, padding="10 10 10 10")
        frame.grid(row=0, column=0, sticky=tk.NSEW)

        ttk.Label(frame, text="Person:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.person_var = tk.StringVar(value=initial_person if initial_person else (global_personnel_list[0] if global_personnel_list else ""))
        self.person_combobox = ttk.Combobox(frame, textvariable=self.person_var, values=global_personnel_list, state='readonly')
        self.person_combobox.grid(row=0, column=1, padx=10, pady=5)
        if initial_person:
            self.person_combobox.set(initial_person)

        ttk.Label(frame, text="Role:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.role_entry = ttk.Entry(frame)
        self.role_entry.grid(row=1, column=1, padx=10, pady=5)
        self.role_entry.insert(0, initial_role)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="OK", command=self.on_ok).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.top.destroy).grid(row=0, column=1, padx=5)

    def on_ok(self):
        pname = self.person_var.get().strip()
        role = self.role_entry.get().strip()
        if not pname:
            messagebox.showerror("Error", "No person selected.", parent=self.top)
            return
        self.result = (pname, role)
        self.top.destroy()


class ManageSubProcessesDialog:
    def __init__(self, parent, subprocesses, global_personnel):
        self.top = tk.Toplevel(parent)
        self.top.title("Manage Sub-Processes")
        self.top.configure(bg=COLORS['background'])
        self.top.transient(parent)
        self.top.grab_set()

        self.subprocesses = subprocesses
        self.global_personnel = global_personnel

        frame = ttk.Frame(self.top, padding="10 10 10 10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Sub-Processes:").grid(row=0, column=0, columnspan=5, sticky=tk.W)
        self.sp_listbox = tk.Listbox(frame, height=10, width=60, bg='white', fg=COLORS['text'],
                                     selectbackground=COLORS['highlight'], font=('Segoe UI', 10))
        self.sp_listbox.grid(row=1, column=0, columnspan=5, pady=(5,10))
        self.update_sp_list()

        add_button = ttk.Button(frame, text="Add", command=self.add_sp)
        add_button.grid(row=2, column=0, sticky=tk.EW, padx=5, pady=5)

        edit_button = ttk.Button(frame, text="Edit", command=self.edit_sp)
        edit_button.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)

        remove_button = ttk.Button(frame, text="Remove", command=self.remove_sp)
        remove_button.grid(row=2, column=2, sticky=tk.EW, padx=5, pady=5)

        personnel_button = ttk.Button(frame, text="Manage Sub-Process Personnel", command=self.manage_sp_personnel)
        personnel_button.grid(row=2, column=3, sticky=tk.EW, padx=5, pady=5)

        close_button = ttk.Button(frame, text="Close", command=self.top.destroy)
        close_button.grid(row=2, column=4, sticky=tk.EW, padx=5, pady=5)

    def update_sp_list(self):
        self.sp_listbox.delete(0, tk.END)
        for sp in self.subprocesses:
            line = f"{sp['name']} ({sp['status']}), Start: {sp['start_date']}, End: {sp['end_date']}"
            self.sp_listbox.insert(tk.END, line)

    def get_selected_sp_index(self):
        sel = self.sp_listbox.curselection()
        if not sel:
            messagebox.showerror("Error", "No sub-process selected.", parent=self.top)
            return None
        return sel[0]

    def add_sp(self):
        dialog = SubProcessDialog(self.top, "Add Sub-Process")
        self.top.wait_window(dialog.top)
        if dialog.result is not None:
            sp_name, sp_status, sp_start, sp_end = dialog.result
            subprocess = {
                "name": sp_name,
                "status": sp_status,
                "start_date": sp_start,
                "end_date": sp_end,
                "personnel": []
            }
            self.subprocesses.append(subprocess)
            self.update_sp_list()

    def edit_sp(self):
        idx = self.get_selected_sp_index()
        if idx is None:
            return
        sp = self.subprocesses[idx]
        dialog = SubProcessDialog(self.top, "Edit Sub-Process",
                                  initial_name=sp["name"],
                                  initial_status=sp["status"],
                                  initial_start_date=sp["start_date"],
                                  initial_end_date=sp["end_date"])
        self.top.wait_window(dialog.top)
        if dialog.result is not None:
            sp_name, sp_status, sp_start, sp_end = dialog.result
            sp["name"] = sp_name
            sp["status"] = sp_status
            sp["start_date"] = sp_start
            sp["end_date"] = sp_end
            self.update_sp_list()

    def remove_sp(self):
        idx = self.get_selected_sp_index()
        if idx is None:
            return
        sp = self.subprocesses[idx]
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the sub-process '{sp['name']}'?", parent=self.top):
            self.subprocesses.pop(idx)
            self.update_sp_list()

    def manage_sp_personnel(self):
        idx = self.get_selected_sp_index()
        if idx is None:
            return
        sp = self.subprocesses[idx]
        dialog = ManageAssignmentsDialog(self.top, sp["personnel"], self.global_personnel, "Sub-Process Personnel")
        self.top.wait_window(dialog.top)
        sp["personnel"] = dialog.assignments
        self.update_sp_list()


class SubProcessDialog:
    def __init__(self, parent, title, initial_name="", initial_status="Not Started", initial_start_date="", initial_end_date=""):
        self.result = None
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.configure(bg=COLORS['background'])
        self.top.transient(parent)
        self.top.grab_set()

        frame = ttk.Frame(self.top, padding="10 10 10 10")
        frame.grid(row=0, column=0, sticky=tk.NSEW)

        ttk.Label(frame, text="Sub-Process Name:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.name_entry = ttk.Entry(frame)
        self.name_entry.grid(row=0, column=1, padx=10, pady=5)
        self.name_entry.insert(0, initial_name)

        ttk.Label(frame, text="Status:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.status_var = tk.StringVar(value=initial_status)
        self.status_combobox = ttk.Combobox(frame, textvariable=self.status_var, values=STATUS_OPTIONS, state='readonly')
        self.status_combobox.grid(row=1, column=1, padx=10, pady=5)
        self.status_combobox.set(initial_status)

        ttk.Label(frame, text="Start Date:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
        self.start_date_entry = ttk.Entry(frame)
        self.start_date_entry.grid(row=2, column=1, padx=10, pady=5)
        self.start_date_entry.insert(0, initial_start_date)

        ttk.Label(frame, text="End Date:").grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
        self.end_date_entry = ttk.Entry(frame)
        self.end_date_entry.grid(row=3, column=1, padx=10, pady=5)
        self.end_date_entry.insert(0, initial_end_date)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="OK", command=self.on_ok).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.top.destroy).grid(row=0, column=1, padx=5)

    def on_ok(self):
        sp_name = self.name_entry.get().strip()
        sp_status = self.status_var.get()
        sp_start = self.start_date_entry.get().strip()
        sp_end = self.end_date_entry.get().strip()

        if not sp_name:
            messagebox.showerror("Error", "Sub-Process name cannot be empty.", parent=self.top)
            return
        if sp_status not in STATUS_OPTIONS:
            messagebox.showerror("Error", "Invalid status.", parent=self.top)
            return

        self.result = (sp_name, sp_status, sp_start, sp_end)
        self.top.destroy()


class GlobalPersonnelDialog(ManagePersonnelDialog):
    pass


def main():
    root = tk.Tk()
    root.geometry("1000x600")
    app = ProjectManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
