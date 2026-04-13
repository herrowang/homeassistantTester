import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import paho.mqtt.client as mqtt
from paho.mqtt.client import CallbackAPIVersion
import json
import threading
from openpyxl import Workbook
from datetime import datetime

class ZWaveMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("Z-Wave JS MQTT Monitor & Stats Collector")
        self.root.geometry("900x700")

        # Data structure for nodes: { node_id: { data } }
        self.nodes_data = {}

        # --- GUI Setup ---
        # 1. Configuration Frame
        setup_frame = ttk.LabelFrame(root, text="Broker Configuration")
        setup_frame.pack(padx=10, pady=5, fill="x")

        # IP and Port
        ttk.Label(setup_frame, text="Broker IP:").grid(row=0, column=0, sticky="w", padx=5)
        self.entry_ip = ttk.Entry(setup_frame)
        self.entry_ip.insert(0, "10.0.0.198")
        self.entry_ip.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(setup_frame, text="Port:").grid(row=0, column=2, sticky="w", padx=5)
        self.entry_port = ttk.Entry(setup_frame, width=8)
        self.entry_port.insert(0, "1883")
        self.entry_port.grid(row=0, column=3, padx=5, pady=2)

        # Credentials
        ttk.Label(setup_frame, text="Username:").grid(row=1, column=0, sticky="w", padx=5)
        self.entry_user = ttk.Entry(setup_frame)
        self.entry_user.insert(0, "vision")
        self.entry_user.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(setup_frame, text="Password:").grid(row=1, column=2, sticky="w", padx=5)
        self.entry_pass = ttk.Entry(setup_frame, show="*")
        self.entry_pass.insert(0, "vision69814136")
        self.entry_pass.grid(row=1, column=3, padx=5, pady=2)

        # Topic
        ttk.Label(setup_frame, text="Base Topic:").grid(row=2, column=0, sticky="w", padx=5)
        self.available_topics = [
            "zwave/#", 
            "zwave/_EVENTS/node/#", 
            "zwave/_STATISTICS/node/#",
            "homeassistant/#"
        ]
        self.entry_topic = ttk.Combobox(setup_frame, values=self.available_topics)
        self.entry_topic.set("zwave/#")
        self.entry_topic.grid(row=2, column=1, columnspan=2, sticky="ew", padx=5, pady=2)
        
        # Bind events for searching
        self.entry_topic.bind("<KeyRelease>", self.on_topic_key_release)
        self.entry_topic.bind("<<ComboboxSelected>>", lambda e: self.log(f"📍 已選擇主題: {self.entry_topic.get()}"))

        self.btn_connect = ttk.Button(setup_frame, text="Connect", command=self.start_mqtt)
        self.btn_connect.grid(row=2, column=3, padx=5, pady=5)

        # 2. Node Table Frame
        table_frame = ttk.LabelFrame(root, text="Z-Wave Nodes Statistics")
        table_frame.pack(padx=10, pady=5, fill="both", expand=True)

        columns = ("home_id", "node_id", "name", "tx", "rx", "dropped_tx", "timeouts", "failure_rate")
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        
        self.tree.heading("home_id", text="Home ID")
        self.tree.heading("node_id", text="Node ID")
        self.tree.heading("name", text="Device Name")
        self.tree.heading("tx", text="TX (Sent)")
        self.tree.heading("rx", text="RX (Recv)")
        self.tree.heading("dropped_tx", text="Dropped TX")
        self.tree.heading("timeouts", text="Timeouts")
        self.tree.heading("failure_rate", text="Failure Rate")

        # Column settings
        col_widths = {"home_id": 80, "node_id": 60, "name": 200, "tx": 80, "rx": 80, "dropped_tx": 80, "timeouts": 80, "failure_rate": 90}
        for col, width in col_widths.items():
            self.tree.column(col, width=width, anchor="center")
        self.tree.column("name", anchor="w")

        self.tree.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # 3. Actions Frame
        actions_frame = tk.Frame(root)
        actions_frame.pack(padx=10, pady=5, fill="x")

        self.btn_export = ttk.Button(actions_frame, text="Export to Excel", command=self.export_excel)
        self.btn_export.pack(side="right", padx=5)

        # 4. Log Area
        self.log_area = scrolledtext.ScrolledText(root, state='disabled', height=10)
        self.log_area.pack(padx=10, pady=10, fill="x")

        # Create Context Menu for Log Area
        self.log_menu = tk.Menu(root, tearoff=0)
        self.log_menu.add_command(label="Copy Selected", command=self.copy_log_selection)
        self.log_menu.add_command(label="Clear Log", command=self.clear_log)
        self.log_area.bind("<Button-3>", self.show_log_menu) # Right-click on Windows/Linux

        # MQTT Client Internal
        self.client = None

    def log(self, message):
        """Thread-safe logging to the UI."""
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def show_log_menu(self, event):
        """Show context menu on right click."""
        self.log_menu.post(event.x_root, event.y_root)

    def copy_log_selection(self):
        """Copy selected text from log area to clipboard."""
        try:
            selected_text = self.log_area.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
        except tk.TclError:
            # No selection
            pass

    def clear_log(self):
        """Clear all text from the log area."""
        self.log_area.configure(state='normal')
        self.log_area.delete('1.0', tk.END)
        self.log_area.configure(state='disabled')

    def on_connect(self, client, userdata, flags, rc, properties=None):
        if rc == 0:
            self.log("✅ 連線成功！")
            topic = self.entry_topic.get()
            client.subscribe(topic)
            self.log(f"📡 已訂閱主題: {topic}")
            self.root.after(0, lambda: self.btn_connect.configure(text="Disconnect"))
        else:
            self.log(f"❌ 連線失敗 (代碼: {rc})")
            self.root.after(0, lambda: self.btn_connect.configure(text="Connect"))

    def on_disconnect(self, client, userdata, flags, rc, properties=None):
        self.log(f"ℹ️ 已斷開連線 (代碼: {rc})")
        self.root.after(0, lambda: self.btn_connect.configure(text="Connect"))

    def on_message(self, client, userdata, msg):
        try:
            topic = msg.topic
            payload_raw = msg.payload.decode()
            payload = json.loads(payload_raw)
            self.log("已收到訊息:"+topic+"["+payload_raw+"]")
            
            # 1. Check for Statistics
            is_stats = topic.endswith("/statistics") or "/statistics/" in topic
            
            if is_stats and isinstance(payload, dict):
                node_id = payload.get("nodeId") or payload.get("id")
                stats_source = payload
                if "data" in payload and isinstance(payload["data"], dict):
                    stats_source = payload["data"]
                    if not node_id:
                        node_id = stats_source.get("nodeId") or stats_source.get("id")

                if not node_id:
                    parts = topic.split("/")
                    if parts[-1].isdigit():
                        node_id = parts[-1]
                    elif len(parts) > 1:
                        node_id = parts[1]

                if node_id:
                    stats = {
                        "tx": stats_source.get("commandsSent", 0),
                        "rx": stats_source.get("commandsReceived", 0),
                        "dropped_tx": stats_source.get("commandsDroppedTX", 0),
                        "timeouts": stats_source.get("timeoutResponse", 0)
                    }
                    self.update_node_data(node_id, stats)
                    return

            # 2. Check for Node Metadata / Status
            parts = topic.split("/")
            if len(parts) >= 2:
                node_label = parts[1]
                if node_label not in ["_CLIENT", "_EVENTS", "_STATISTICS"]:
                    node_info = {}
                    if topic.endswith("/status") and isinstance(payload, dict):
                        node_info["name"] = payload.get("name")
                        node_info["home_id"] = payload.get("homeid")
                        node_info["node_id"] = payload.get("id")
                    
                    self.update_node_data(node_label, node_info)

        except Exception as e:
            pass

    def on_topic_key_release(self, event):
        """Filter the combobox values based on user input."""
        value = event.widget.get().lower()
        if value == '':
            self.entry_topic['values'] = self.available_topics
        else:
            data = []
            for item in self.available_topics:
                if value in item.lower():
                    data.append(item)
            self.entry_topic['values'] = data
        
        self.entry_topic.event_generate('<Down>')

    def update_node_data(self, node_id, new_data):
        """Thread-safe update of node information and UI tree."""
        node_id = str(node_id)
        if node_id not in self.nodes_data:
            self.nodes_data[node_id] = {
                "home_id": "", "node_id": node_id, "name": f"Node {node_id}",
                "tx": 0, "rx": 0, "dropped_tx": 0, "timeouts": 0, "failure_rate": "0.0%"
            }
        
        self.nodes_data[node_id].update(new_data)
        
        node_name = self.nodes_data[node_id].get("id", node_id)
        new_topic = f"zwave/{node_name}/#"
        if new_topic not in self.available_topics:
            self.available_topics.append(new_topic)
            self.root.after(0, lambda: self.entry_topic.configure(values=self.available_topics))

        d = self.nodes_data[node_id]
        total_attempts = d["tx"] + d["dropped_tx"] + d["timeouts"]
        failures = d["dropped_tx"] + d["timeouts"]
        
        if total_attempts > 0:
            rate = (failures / total_attempts) * 100
            d["failure_rate"] = f"{rate:.1f}%"
        else:
            d["failure_rate"] = "0.0%"

        self.root.after(0, self.refresh_tree)

    def refresh_tree(self):
        """Update the Treeview with current nodes_data."""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for node_id in sorted(self.nodes_data.keys(), key=lambda x: int(x) if x.isdigit() else 0):
            d = self.nodes_data[node_id]
            self.tree.insert("", tk.END, values=(
                d["home_id"], d["node_id"], d["name"], 
                d["tx"], d["rx"], d["dropped_tx"], d["timeouts"], d["failure_rate"]
            ))

    def export_excel(self):
        """Export the collected node data to an Excel file."""
        if not self.nodes_data:
            messagebox.showwarning("Warning", "No data to export.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"ZWave_Stats_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Z-Wave Statistics"
            headers = ["Home ID", "Node ID", "Device Name", "TX (Sent)", "RX (Received)", "Dropped TX", "Timeouts", "Failure Rate"]
            ws.append(headers)
            for node_id in sorted(self.nodes_data.keys(), key=lambda x: int(x) if x.isdigit() else 0):
                d = self.nodes_data[node_id]
                ws.append([
                    d["home_id"], d["node_id"], d["name"], 
                    d["tx"], d["rx"], d["dropped_tx"], d["timeouts"], d["failure_rate"]
                ])
            wb.save(file_path)
            messagebox.showinfo("Success", f"Data exported successfully to:\n{file_path}")
            self.log(f"📁 Exported stats to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export Excel: {e}")

    def start_mqtt(self):
        if self.client and self.client.is_connected():
            self.stop_mqtt()
            return

        ip = self.entry_ip.get()
        port = int(self.entry_port.get())
        user = self.entry_user.get()
        pw = self.entry_pass.get()

        self.client = mqtt.Client(CallbackAPIVersion.VERSION2)
        self.client.on_connect = self.on_connect
        self.client.on_message = self.on_message
        self.client.on_disconnect = self.on_disconnect

        if user:
            self.client.username_pw_set(user, pw)

        try:
            self.client.connect_async(ip, port, 60)
            self.client.loop_start()
            self.log(f"🔄 Connecting to {ip}...")
            self.btn_connect.configure(text="Connecting...")
        except Exception as e:
            messagebox.showerror("Error", f"Could not connect: {e}")
            self.btn_connect.configure(text="Connect")

    def stop_mqtt(self):
        if self.client:
            self.client.loop_stop()
            self.client.disconnect()
            self.log("🔌 已主動斷開連線")
            self.btn_connect.configure(text="Connect")

if __name__ == "__main__":
    root = tk.Tk()
    app = ZWaveMonitor(root)
    root.mainloop()
