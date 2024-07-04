import platform
import socket
import psutil
import pandas as pd
import subprocess
import datetime
from getmac import get_mac_address as gma
import tkinter as tk
from tkinter import filedialog

from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


# Function to get system details
def get_system_details():
    system_details = {
        "Device Name": platform.node(),
        "Processor": platform.processor(),
        "Installed RAM": f"{round(psutil.virtual_memory().total / (1024**3), 2)} GB",
        "System Type": platform.machine(),
        "Edition": platform.system(),
        "Version": platform.version(),
    }
    return system_details

# Function to get logged-in users
def get_users():
    users = psutil.users()
    user_details = []
    for user in users:
        user_details.append({
            "User Name": user.name,
            "Terminal": user.terminal,
            "Host": user.host,
            "Started": datetime.datetime.fromtimestamp(user.started).strftime("%Y-%m-%d %H:%M:%S")
        })
    return user_details

# Function to get MAC addresses
def get_mac_addresses():
    mac_address = gma()
    return [{"MAC Address": mac_address}]

# Function to get IP addresses
def get_ip_addresses():
    interfaces = psutil.net_if_addrs()
    ip_addresses = []
    for interface_name, addresses in interfaces.items():
        for addr in addresses:
            if addr.family == socket.AF_INET:
                ip_addresses.append({
                    "Interface": interface_name,
                    "IP Address": addr.address,
                    "Netmask": addr.netmask,
                    "Broadcast IP": addr.broadcast
                })
    return ip_addresses

# Function to get installed applications (Windows example)
def get_installed_apps_windows():
    apps = []
    cmd = 'powershell "Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate"'
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
    lines = result.stdout.split('\n')[3:]
    for line in lines:
        if line.strip():
            columns = [x.strip() for x in line.split(None, 3)]
            app = {
                "Name": columns[0] if len(columns) > 0 else "",
                "Version": columns[1] if len(columns) > 1 else "",
                "Publisher": columns[2] if len(columns) > 2 else "",
                "InstallDate": columns[3] if len(columns) > 3 else ""
            }
            apps.append(app)
    return apps

# Function to create an Excel report
def create_excel_report(system_details, user_details, mac_addresses, ip_addresses, installed_apps, save_path):
    excel_filename = save_path + '/system_audit_report.xlsx'
    with pd.ExcelWriter(excel_filename) as writer:
        # Write system details
        system_df = pd.DataFrame([system_details])
        system_df.to_excel(writer, sheet_name='System Details', index=False)
        
        # Write user details
        users_df = pd.DataFrame(user_details)
        users_df.to_excel(writer, sheet_name='System Users', index=False)
        
        # Write MAC addresses
        mac_df = pd.DataFrame(mac_addresses)
        mac_df.to_excel(writer, sheet_name='MAC Addresses', index=False)
        
        # Write IP addresses
        ip_df = pd.DataFrame(ip_addresses)
        ip_df.to_excel(writer, sheet_name='IP Addresses', index=False)
        
        # Write installed applications
        apps_df = pd.DataFrame(installed_apps)
        apps_df.to_excel(writer, sheet_name='Installed Applications', index=False)

# Function to create PDF report
def create_pdf_report(system_details, user_details, mac_address, ip_addresses, installed_apps, save_path):
    pdf_filename = save_path + '/system_audit_report.pdf'
    doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(letter))
    
    # Container for the 'Flowable' objects
    elements = []
    
    # Title
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    title = Paragraph("System Audit Report", title_style)
    elements.append(title)
    
    # Date and Time
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    date_time = Paragraph(f"Report generated on: {now}", styles['Normal'])
    elements.append(date_time)
    
    # System Details
    system_table_data = [[key, value] for key, value in system_details.items()]
    system_table = Table(system_table_data, colWidths=[200, 300])
    system_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                                      ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                      ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                      ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
    elements.append(Paragraph("System Details:", styles['Heading2']))
    elements.append(system_table)
    
    # User Details
    user_table_data = [["User Name", "Terminal", "Host", "Started"]] + \
                      [[user['User Name'], user['Terminal'], user['Host'], user['Started']] for user in user_details]
    user_table = Table(user_table_data, colWidths=[100, 100, 100, 200])
    user_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                    ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
    elements.append(Paragraph("Logged-in Users:", styles['Heading2']))
    elements.append(user_table)
    
    # MAC Address
    elements.append(Paragraph(f"MAC Address: {mac_address}", styles['Normal']))
    
    # IP Addresses
    ip_table_data = [["Interface", "IP Address", "Netmask", "Broadcast IP"]] + \
                    [[ip['Interface'], ip['IP Address'], ip['Netmask'], ip['Broadcast IP']] for ip in ip_addresses]
    ip_table = Table(ip_table_data, colWidths=[100, 100, 100, 100])
    ip_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                                  ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                  ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                  ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
    elements.append(Paragraph("IP Addresses:", styles['Heading2']))
    elements.append(ip_table)
    
    # Installed Applications
    if installed_apps:
        app_table_data = [["Name", "Version", "Publisher", "Install Date"]] + \
                         [[app['Name'], app['Version'], app['Publisher'], app['InstallDate']] for app in installed_apps]
        app_table = Table(app_table_data, colWidths=[200, 100, 150, 100])
        app_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                                       ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                       ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                       ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
        elements.append(Paragraph("Installed Applications:", styles['Heading2']))
        elements.append(app_table)
    else:
        elements.append(Paragraph("Installed Applications: None", styles['Heading2']))
    
    # Build the PDF document
    doc.build(elements)
    print(f"PDF report generated successfully: {pdf_filename}")

def get_save_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    save_path = filedialog.askdirectory(title="Select Directory to Save Reports")
    root.destroy()  # Destroy the Tkinter window
    
    return save_path

if __name__ == "__main__":
    save_path = get_save_path()
    if save_path:
        system_details = get_system_details()
        user_details = get_users()
        mac_addresses = get_mac_addresses()
        ip_addresses = get_ip_addresses()
        installed_apps = get_installed_apps_windows() if platform.system() == "Windows" else []  # Adjust for other OS
        
        create_excel_report(system_details, user_details, mac_addresses, ip_addresses, installed_apps, save_path)
        create_pdf_report(system_details, user_details, mac_addresses, ip_addresses, installed_apps, save_path)
        print("System audit reports generated successfully!")
    else:
        print("No directory selected. Reports were not generated.")
