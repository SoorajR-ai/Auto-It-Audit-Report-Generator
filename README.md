<!DOCTYPE html>
<html lang="en">
<body>
    <h1>System Audit Report Generator</h1>
    <p>This script collects various system details such as device information, logged-in users, MAC addresses, IP addresses, and installed applications. It then generates both an Excel and a PDF report containing this information.</p>
    <h2>Features</h2>
    <ul>
        <li>Collects system details (device name, processor, RAM, etc.)</li>
        <li>Retrieves information about logged-in users</li>
        <li>Gets the MAC address and IP addresses of the system</li>
        <li>Gathers a list of installed applications (Windows only)</li>
        <li>Creates Excel and PDF reports</li>
    </ul>
    <h2>Requirements</h2>
    <ul>
        <li>Python 3.x</li>
        <li>Required packages: <code>platform</code>, <code>socket</code>, <code>psutil</code>, <code>pandas</code>, <code>subprocess</code>, <code>datetime</code>, <code>getmac</code>, <code>tkinter</code>, <code>reportlab</code></li>
    </ul>
    <h2>Installation</h2>
    <p>Install the required packages using pip:</p>
    <pre>
        <code>pip install -r requirements.txt</code>
    </pre>
    <h2>Usage</h2>
    <ol>
        <li>Run the script:</li>
        <pre>
            <code>python main.py</code>
        </pre>
        <li>Select the directory where you want to save the reports.</li>
    </ol>
    <p>The script will generate <code>system_audit_report.xlsx</code> and <code>system_audit_report.pdf</code> in the selected directory.</p>
    <h2>Script Details</h2>
    <p>The script includes the following functions:</p>
    <ul>
        <li><code>get_system_details()</code>: Collects system-related information.</li>
        <li><code>get_users()</code>: Retrieves information about logged-in users.</li>
        <li><code>get_mac_addresses()</code>: Gets the MAC address of the system.</li>
        <li><code>get_ip_addresses()</code>: Gathers IP address details for all network interfaces.</li>
        <li><code>get_installed_apps_windows()</code>: Uses a PowerShell command to get a list of installed applications on Windows.</li>
        <li><code>create_excel_report()</code>: Generates an Excel report and saves it to the specified path.</li>
        <li><code>create_pdf_report()</code>: Generates a PDF report and saves it to the specified path.</li>
        <li><code>get_save_path()</code>: Opens a file dialog for the user to select a directory to save the reports.</li>
    </ul>
    <h2>Contributing</h2>
    <p>Contributions are welcome! Please fork this repository and submit a pull request with your changes.</p>
    <h2>License</h2>
    <p>This project is licensed under the MIT License. See the <code>LICENSE</code> file for more details.</p>
</body>
</html>
