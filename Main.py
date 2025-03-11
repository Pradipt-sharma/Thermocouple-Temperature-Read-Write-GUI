# Thermocouple-Temperature-Read-Write-GUI
import tkinter as tk
from tkinter import filedialog, ttk
import serial
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import threading
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import time
import serial.tools.list_ports

# Serial port configuration
BAUD_RATE = 115200
ser = None
logging_active = False
saving_active = True
start_time = None
file_path = None

# Data storage
data_list = {"Seconds": [], "Sensor1": [], "Sensor2": [], "Sensor3": [], "Sensor4": []}
latest_values = {"Sensor1": 0, "Sensor2": 0, "Sensor3": 0, "Sensor4": 0}

# Function to detect available COM ports
def get_com_ports():
    return [port.device for port in serial.tools.list_ports.comports()]

# Function to refresh COM port list
def refresh_com_ports():
    com_ports = get_com_ports()
    com_port_dropdown["values"] = com_ports
    if com_ports:
        com_port_var.set(com_ports[0])  # Select the first available port

# Function to connect to selected COM port
def connect_serial():
    global ser
    selected_port = com_port_var.get()
    try:
        ser = serial.Serial(selected_port, BAUD_RATE, timeout=1)
        connection_status.config(text="Connected", fg="green")
    except Exception as e:
        connection_status.config(text="Disconnected", fg="red")
        print(f"Connection error: {e}")

# Function to read serial data
def read_serial():
    global logging_active, start_time
    start_time = time.time()
    while logging_active and ser and ser.is_open:
        try:
            line = ser.readline().decode('utf-8').strip()
            if line:
                values = line.split(',')
                if len(values) == 4:
                    seconds = round(time.time() - start_time, 2)
                    data_list["Seconds"].append(seconds)
                    for i, key in enumerate(list(data_list.keys())[1:]):
                        val = float(values[i])
                        data_list[key].append(val)
                        latest_values[key] = val
                    if saving_active and file_path:
                        save_continuous()
        except Exception as e:
            print(f"Error reading serial data: {e}")

# Function to save data continuously to Excel
def save_continuous():
    if not saving_active:
        return
    df = pd.DataFrame(data_list)
    df.to_excel(file_path, index=False)
    print("Data updated in Excel.")

# Function to start saving to Excel
def start_saving():
    global file_path, saving_active
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        saving_active = True
        print(f"Saving data to {file_path}")

# Function to update live plot
def update_plot(frame):
    ax.clear()
    for sensor, values in list(data_list.items())[1:]:
        ax.plot(data_list["Seconds"][-50:], values[-50:], label=sensor)
    ax.legend()
    ax.set_xlabel("Time (s)")
    ax.set_ylabel("Temperature (°C)")
    ax.set_title("Live Temperature Data")
    ax.tick_params(axis='x', rotation=45)
    plt.tight_layout()

# Function to update sensor readings display
def update_sensor_labels():
    text = "\n".join([f"{key}: {latest_values[key]:.2f} °C" for key in latest_values])
    sensor_label.config(text=text)
    root.after(500, update_sensor_labels)

# Function to start logging
def start_logging():
    global logging_active
    logging_active = True
    thread = threading.Thread(target=read_serial, daemon=True)
    thread.start()
    print("Logging started...")

# Function to stop logging
def stop_logging():
    global logging_active
    logging_active = False
    print("Logging stopped.")

# Function to stop saving
def stop_saving():
    global saving_active
    saving_active = False
    print("Saving stopped.")

# GUI setup
root = tk.Tk()
root.title("Temperature Data Logger")
root.geometry("1000x700")  # Adjusted for better fullscreen handling

frame = tk.Frame(root)
frame.pack()

com_port_var = tk.StringVar()
com_port_dropdown = ttk.Combobox(frame, textvariable=com_port_var, values=get_com_ports())
com_port_dropdown.grid(row=0, column=0, padx=5, pady=5)
connect_button = tk.Button(frame, text="Connect", command=connect_serial)
connect_button.grid(row=0, column=1, padx=5, pady=5)
connection_status = tk.Label(frame, text="Disconnected", fg="red")
connection_status.grid(row=0, column=2, padx=5, pady=5)
refresh_button = tk.Button(frame, text="Refresh", command=refresh_com_ports)
refresh_button.grid(row=0, column=3, padx=5, pady=5)

start_button = tk.Button(frame, text="Start Logging", command=start_logging)
start_button.grid(row=1, column=0, padx=5, pady=5)
stop_button = tk.Button(frame, text="Stop Logging", command=stop_logging)
stop_button.grid(row=1, column=1, padx=5, pady=5)
save_button = tk.Button(frame, text="Start Saving", command=start_saving)
save_button.grid(row=1, column=2, padx=5, pady=5)
stop_saving_button = tk.Button(frame, text="Stop Saving", command=stop_saving)
stop_saving_button.grid(row=1, column=3, padx=5, pady=5)

# Live plot setup
fig, ax = plt.subplots(figsize=(9, 5))  # Adjusted size
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)  # Adjust to fit full screen
ani = FuncAnimation(fig, update_plot, interval=500)

# Sensor readings display
sensor_label = tk.Label(root, text="", font=("Arial", 12), justify="left")
sensor_label.pack(pady=5)

update_sensor_labels()

root.mainloop()

if ser:
    ser.close()
