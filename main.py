import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import paho.mqtt.client as mqtt
from paho.mqtt.client import CallbackAPIVersion
import json
import threading
from openpyxl import Workbook
from datetime import datetime
import os

class ZWaveMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("Z-Wave JS MQTT Monitor & Stats Collector")
        self.root.geometry("1000x750")

        # Data structure for nodes: { node_id: { data } }
        self.nodes_data = {}
        self.log_file_path = None

        # --- GUI Setup ---
        # Top Frame for both Config and Controller RF Status
        top_container = tk.Frame(root)
        top_container.pack(padx=10, pady=5, fill="x")

        # 1. Configuration Frame
        setup_frame = ttk.LabelFrame(top_container, text="Broker Configuration")
        setup_frame.pack(side="left", padx=(0, 10), fill="both", expand=True)

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
            "homeassistant/#",
            "Mosquitto/node/#",
            "Mosquitto/node/statistics_updated"
        ]
        self.entry_topic = ttk.Combobox(setup_frame, values=self.available_topics)
        self.entry_topic.set("zwave/#")
        self.entry_topic.grid(row=2, column=1, columnspan=2, sticky="ew", padx=5, pady=2)
        
        # Bind events for searching
        self.entry_topic.bind("<KeyRelease>", self.on_topic_key_release)
        self.entry_topic.bind("<<ComboboxSelected>>", lambda e: self.log(f"📍 已選擇主題: {self.entry_topic.get()}"))

        self.btn_connect = ttk.Button(setup_frame, text="Connect", command=self.start_mqtt)
        self.btn_connect.grid(row=2, column=3, padx=5, pady=5)

        # 1.1 Controller RF Status Frame (Top Right)
        rf_frame = ttk.LabelFrame(top_container, text="Controller RF Status")
        rf_frame.pack(side="right", fill="both")

        self.rf_labels = {}
        rf_fields = [
            ("RX", "RX: 0"), ("TX", "TX: 0"),
            ("DroppedRX", "DroppedRX: 0"), ("DroppedTX", "DroppedTX: 0"),
            ("RSSIChannel0", "Ch0: N/A"), ("RSSIChannel1", "Ch1: N/A"),
            ("RSSIChannel2", "Ch2: N/A"), ("RSSIChannel3", "Ch3: N/A")
        ]
        for i, (key, text) in enumerate(rf_fields):
            lbl = ttk.Label(rf_frame, text=text, width=15)
            lbl.grid(row=i//2, column=i%2, padx=5, pady=2, sticky="w")
            self.rf_labels[key] = lbl

        # 2. Node Table Frame
        table_frame = ttk.LabelFrame(root, text="Z-Wave Nodes Statistics")
        table_frame.pack(padx=10, pady=5, fill="both", expand=True)

        columns = ("home_id", "node_id", "name", "tx", "rx", "dropped_tx", "dropped_rx", "timeouts", "rssi", "failure_rate")
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        
        self.tree.heading("home_id", text="Home ID")
        self.tree.heading("node_id", text="Node ID")
        self.tree.heading("name", text="Device Name")
        self.tree.heading("tx", text="TX (Sent)")
        self.tree.heading("rx", text="RX (Recv)")
        self.tree.heading("dropped_tx", text="Dropped TX")
        self.tree.heading("dropped_rx", text="Dropped RX")
        self.tree.heading("timeouts", text="Timeouts")
        self.tree.heading("rssi", text="RSSI")
        self.tree.heading("failure_rate", text="Failure Rate")

        # Column settings
        col_widths = {
            "home_id": 80, "node_id": 60, "name": 180, 
            "tx": 80, "rx": 80, "dropped_tx": 85, "dropped_rx": 85, 
            "timeouts": 80, "rssi": 60, "failure_rate": 90
        }
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

        self.btn_file_log = ttk.Button(actions_frame, text="Start Logging to File", command=self.toggle_file_logging)
        self.btn_file_log.pack(side="right", padx=5)

        # 4. Log Area
        self.log_area = scrolledtext.ScrolledText(root, state='disabled', height=10)
        self.log_area.pack(padx=10, pady=10, fill="x")

        # Create Context Menu for Log Area
        self.log_menu = tk.Menu(root, tearoff=0)
        self.log_menu.add_command(label="Copy Selected", command=self.copy_log_selection)
        self.log_menu.add_command(label="Clear Log", command=self.clear_log)
        self.log_area.bind("<Button-3>", self.show_log_menu) 

        # MQTT Client Internal
        self.client = None
        
        # Start the polling loop
        self.poll_statistics()

    def log(self, message):
        """Thread-safe logging to the UI and optional file."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        full_msg = f"[{timestamp}] {message}"
        
        # UI Log
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, f"{full_msg}\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

        # File Log
        if self.log_file_path:
            try:
                with open(self.log_file_path, "a", encoding="utf-8") as f:
                    f.write(f"{full_msg}\n")
            except Exception:
                pass

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
            pass

    def clear_log(self):
        """Clear all text from the log area."""
        self.log_area.configure(state='normal')
        self.log_area.delete('1.0', tk.END)
        self.log_area.configure(state='disabled')

    def toggle_file_logging(self):
        """Enable or disable real-time logging to a text file."""
        if self.log_file_path:
            self.log(f"📝 已停止記錄日誌至檔案: {os.path.basename(self.log_file_path)}")
            self.log_file_path = None
            self.btn_file_log.configure(text="Start Logging to File")
        else:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfile=f"MQTT_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )
            if file_path:
                self.log_file_path = file_path
                self.btn_file_log.configure(text="Stop Logging to File")
                self.log(f"📝 開始記錄日誌至檔案: {file_path}")

    def poll_statistics(self):
        """Actively request statistics from Z-Wave JS every second."""
        if self.client and self.client.is_connected():
            try:
                command = {"api": "getNodes", "args": []}
                self.client.publish("zwave/_CLIENT/COMMAND/api", json.dumps(command))
            except Exception:
                pass
        self.root.after(1000, self.poll_statistics)

    def on_connect(self, client, userdata, flags, rc, properties=None):
        if rc == 0:
            self.log("✅ 連線成功！")
            topic = self.entry_topic.get()
            client.subscribe(topic)
            client.subscribe("zwave/_CLIENT/REPLY/api")
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
            
            # 1. Handle API Reply from polling (getNodes)
            if topic == "zwave/_CLIENT/REPLY/api" and isinstance(payload, dict):
                if payload.get("success") and "result" in payload:
                    nodes = payload["result"]
                    for node in nodes:
                        node_id = node.get("id")
                        stats_source = node.get("statistics")
                        if node_id and stats_source:
                            stats = {
                                "tx": stats_source.get("commandsTX", stats_source.get("commandsSent")),
                                "rx": stats_source.get("commandsRX", stats_source.get("commandsReceived")),
                                "dropped_tx": stats_source.get("commandsDroppedTX"),
                                "dropped_rx": stats_source.get("commandsDroppedRX"),
                                "timeouts": stats_source.get("timeoutResponse")
                            }
                            node_info = {
                                "name": node.get("name") or node.get("label"),
                                "home_id": node.get("homeid")
                            }
                            self.update_node_data(node_id, {**node_info, **stats})
                    return

            # 2. Handle nodeinfo / status topics (Optimized for the provided example)
            if topic.endswith("/nodeinfo") or topic.endswith("/status"):
                nodes_to_process = []
                if isinstance(payload, list):
                    nodes_to_process = payload
                elif isinstance(payload, dict):
                    nodes_to_process = [payload]
                
                for node_data in nodes_to_process:
                    node_id = node_data.get("id")
                    if node_id is not None:
                        # Extract from statistics object as per example
                        stats_box = node_data.get("statistics", {})
                        if stats_box: # 只有在 statistics 有內容時才處理
                            extracted_data = {
                                "tx": stats_box.get("commandsTX", stats_box.get("commandsSent")),
                                "rx": stats_box.get("commandsRX", stats_box.get("commandsReceived")),
                                "dropped_tx": stats_box.get("commandsDroppedTX"),
                                "dropped_rx": stats_box.get("commandsDroppedRX"),
                                "timeouts": stats_box.get("timeoutResponse"),
                                "name": node_data.get("name") or node_data.get("label"),
                                "home_id": node_data.get("homeid")
                            }
                            self.update_node_data(node_id, extracted_data)
                
                if not topic.startswith("zwave/_CLIENT/REPLY/api"):
                    self.log(f"📋 已更新節點統計 (自 {topic.split('/')[-1]})")
                return

            # 2.1 Handle Mosquitto/node/statistics_updated
            if topic.startswith("zwave/_EVENTS/ZWAVE_GATEWAY-Mosquitto/node/statistics_updated") and isinstance(payload, dict):
                data_list = payload.get("data")
                if len(data_list) >= 2:
                    node_info = data_list[0]
                    stats_source = data_list[1]
                    node_id = node_info.get("id")
                    if node_id is not None:
                        extracted_data = {
                            "name": node_info.get("name"),
                            "tx": stats_source.get("commandsTX"),
                            "rx": stats_source.get("commandsRX"),
                            "dropped_tx": stats_source.get("commandsDroppedTX"),
                            "dropped_rx": stats_source.get("commandsDroppedRX"),
                            "timeouts": stats_source.get("timeoutResponse"),
                            "rssi": stats_source.get("rssi")
                        }
                        self.update_node_data(node_id, extracted_data)
                self.log("📋 已更新節點統計 (自 statistics_updated)")
                return
            if topic.startswith("zwave/_EVENTS/ZWAVE_GATEWAY-Mosquitto/controller/statistics_updated") and isinstance(payload, dict):
                controller_data = payload.get("data")
                if isinstance(controller_data, list) and len(controller_data) > 0:
                    stats = controller_data[0]
                    # Get background RSSI if available
                    bg_rssi = stats.get("backgroundRSSI", {})
                    extracted_rf_data = {
                        "RX": stats.get("messagesRX", 0),
                        "TX": stats.get("messagesTX", 0),
                        "DroppedRX": stats.get("messagesDroppedRX", 0),
                        "DroppedTX": stats.get("messagesDroppedTX", 0),
                        "RSSIChannel0": bg_rssi.get("channel0", "N/A"),
                        "RSSIChannel1": bg_rssi.get("channel1", "N/A"),
                        "RSSIChannel2": bg_rssi.get("channel2", "N/A"),
                        "RSSIChannel3": bg_rssi.get("channel3", "N/A"),
                    }
                    self.root.after(0, lambda: self.update_rf_status(extracted_rf_data))
                self.log("📋 已更新控制器統計 (自 controller/statistics_updated)")
                return
            # Log other general messages
            if not topic.startswith("zwave/_CLIENT/REPLY/api"):
                self.log("已收到訊息:"+topic+"["+payload_raw+"]")

            # 3. Check for Statistics (Event-based messages)
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
                    if parts[-1].isdigit(): node_id = parts[-1]
                    elif len(parts) > 1: node_id = parts[1]

                if node_id:
                    stats = {
                        "tx": stats_source.get("commandsTX", stats_source.get("commandsSent")),
                        "rx": stats_source.get("commandsRX", stats_source.get("commandsReceived")),
                        "dropped_tx": stats_source.get("commandsDroppedTX"),
                        "dropped_rx": stats_source.get("commandsDroppedRX"),
                        "timeouts": stats_source.get("timeoutResponse"),
                        "rssi": stats_source.get("rssi")
                    }
                    self.update_node_data(node_id, stats)
                    return

        except Exception:
            pass

    def on_topic_key_release(self, event):
        value = event.widget.get().lower()
        if value == '':
            self.entry_topic['values'] = self.available_topics
        else:
            data = [item for item in self.available_topics if value in item.lower()]
            self.entry_topic['values'] = data
        self.entry_topic.event_generate('<Down>')

    def update_rf_status(self, data):
        """Update the Controller RF Status labels."""
        mapping = {
            "RX": f"RX: {data.get('RX')}",
            "TX": f"TX: {data.get('TX')}",
            "DroppedRX": f"DroppedRX: {data.get('DroppedRX')}",
            "DroppedTX": f"DroppedTX: {data.get('DroppedTX')}",
            "RSSIChannel0": f"Ch0: {data.get('RSSIChannel0')} dBm",
            "RSSIChannel1": f"Ch1: {data.get('RSSIChannel1')} dBm",
            "RSSIChannel2": f"Ch2: {data.get('RSSIChannel2')} dBm",
            "RSSIChannel3": f"Ch3: {data.get('RSSIChannel3')} dBm"
        }
        for key, text in mapping.items():
            if key in self.rf_labels:
                self.rf_labels[key].configure(text=text)

    def update_node_data(self, node_id, new_data):
        node_id = str(node_id)
        if node_id not in self.nodes_data:
            self.nodes_data[node_id] = {
                "home_id": "", "node_id": node_id, "name": f"Node {node_id}",
                "tx": 0, "rx": 0, "dropped_tx": 0, "dropped_rx": 0, "timeouts": 0, "rssi": "N/A", "failure_rate": "0.0%"
            }
        
        # Only update if the value is not None/Zero or if it's a fresh update
        for k, v in new_data.items():
            if v is not None:
                self.nodes_data[node_id][k] = v
        
        # Discover new topic based on name
        node_name = self.nodes_data[node_id].get("name", node_id)
        new_topic = f"zwave/{node_name}/#"
        if new_topic not in self.available_topics:
            self.available_topics.append(new_topic)
            self.root.after(0, lambda: self.entry_topic.configure(values=self.available_topics))

        # Calculate failure rate
        d = self.nodes_data[node_id]
        total_attempts = (d["tx"] or 0) + (d["dropped_tx"] or 0) + (d["timeouts"] or 0)
        failures = (d["dropped_tx"] or 0) + (d["timeouts"] or 0)
        
        if total_attempts > 0:
            rate = (failures / total_attempts) * 100
            d["failure_rate"] = f"{rate:.1f}%"
        else:
            d["failure_rate"] = "0.0%"
        self.root.after(0, self.refresh_tree)

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for node_id in sorted(self.nodes_data.keys(), key=lambda x: int(x) if x.isdigit() else 0):
            d = self.nodes_data[node_id]
            self.tree.insert("", tk.END, values=(
                d["home_id"], d["node_id"], d["name"], 
                d["tx"], d["rx"], d["dropped_tx"], d["dropped_rx"], d["timeouts"], d["rssi"], d["failure_rate"]
            ))

    def export_excel(self):
        if not self.nodes_data:
            messagebox.showwarning("Warning", "No data to export.")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"ZWave_Stats_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not file_path: return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Z-Wave Statistics"
            headers = ["Home ID", "Node ID", "Device Name", "TX (Sent)", "RX (Received)", "Dropped TX", "Dropped RX", "Timeouts", "RSSI", "Failure Rate"]
            ws.append(headers)
            for node_id in sorted(self.nodes_data.keys(), key=lambda x: int(x) if x.isdigit() else 0):
                d = self.nodes_data[node_id]
                ws.append([
                    d["home_id"], d["node_id"], d["name"], 
                    d["tx"], d["rx"], d["dropped_tx"], d["dropped_rx"], d["timeouts"], d["rssi"], d["failure_rate"]
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
        if user: self.client.username_pw_set(user, pw)
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
