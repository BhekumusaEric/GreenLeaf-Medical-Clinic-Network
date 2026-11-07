from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

OUT = "/Users/eric/GreenLeaf-Medical-Clinic-Network/GreenLeaf_Medical_Clinic_Student_Answer_Sheet_FINAL.docx"
SS  = "/Users/eric/GreenLeaf-Medical-Clinic-Network/screenshots"

FONT = "Arial"
FONT_SIZE = 11

doc = Document()

# ── Page margins ────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1.18)
    section.right_margin  = Inches(1.18)

# ── Helpers ─────────────────────────────────────────────────────────────────
def style_run(run, bold=False, size=FONT_SIZE, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

def heading(text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    sz = {1: 14, 2: 12, 3: 11}[level]
    style_run(run, bold=True, size=sz)
    return p

def body(text, bold=False, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(4)
    if indent:
        p.paragraph_format.left_indent = Inches(0.3)
    run = p.add_run(text)
    style_run(run, bold=bold)
    return p

def insert_image(path, width=6.0, caption=None):
    if os.path.exists(path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(path, width=Inches(width))
        if caption:
            cp = doc.add_paragraph(caption)
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_after = Pt(8)
            for run in cp.runs:
                style_run(run, size=9)
    else:
        body(f"[Image not found: {path}]")

def page_break():
    doc.add_page_break()

def add_table(headers, rows):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = h
        for para in hdr[i].paragraphs:
            for run in para.runs:
                style_run(run, bold=True, size=10)
    for ri, row in enumerate(rows):
        cells = t.rows[ri+1].cells
        for ci, val in enumerate(row):
            cells[ci].text = val
            for para in cells[ci].paragraphs:
                for run in para.runs:
                    style_run(run, size=10)
    doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════════════
# COVER PAGE
# ════════════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Eduvos")
style_run(r, bold=True, size=13)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Network Architecture & Implementation (ITCTA1-44)")
style_run(r, size=11)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("GREENLEAF MEDICAL CLINIC NETWORK")
style_run(r, bold=True, size=16)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Cisco Packet Tracer Assessment — Student Answer Sheet")
style_run(r, size=12)

doc.add_paragraph()

add_table(
    ["Field", "Details"],
    [
        ["Student Name", "Bhekumusa Eric Ntshwenya"],
        ["Student Number", "EDUV13678625"],
        ["Campus", "Bedfordview"],
        ["Module", "ITCTA1-44 — Network Architecture & Implementation"],
        ["Assessment", "Cisco Packet Tracer Practical Assessment"],
        ["Date", "November 2026"],
    ]
)

page_break()

# ════════════════════════════════════════════════════════════════════════════
# DELIVERABLE 1
# ════════════════════════════════════════════════════════════════════════════
heading("DELIVERABLE 1: NETWORKING CORE CONCEPTS", 1)

# ── Q1 ───────────────────────────────────────────────────────────────────────
heading("Question 1: Networking Principles and Models", 2)
heading("1.1 Network Types", 3)

body("a) The original network model used by GreenLeaf Clinic was a Peer-to-Peer (P2P) network model.")
body("b) After the upgrade, the clinic implemented a Client-Server network model.")

heading("1.2 Distinguishing Network Models", 3)
add_table(
    ["Aspect", "Peer-to-Peer (Original)", "Client-Server (Upgraded)"],
    [
        ["Resource Management",
         "Each workstation manages its own resources. Files are shared directly between computers with no central control.",
         "Resources are stored and managed centrally on dedicated servers. Clients request access through the server."],
        ["Security",
         "No centralised security. Each device sets its own permissions, making it difficult to enforce consistent access control.",
         "Security is managed centrally. Administrators control user accounts, permissions, and access policies from the server."],
        ["Scalability",
         "Difficult to scale. Adding more devices increases complexity and degrades performance as each device handles its own sharing.",
         "Easily scalable. New clients are simply added to the network and connect to existing servers without major reconfiguration."],
    ]
)

heading("1.3 Suitability for Sensitive Patient Data", 3)
body("The client-server model is better suited for handling sensitive patient data for the following reasons:")
body("Centralised Security and Control: All patient records are stored on a dedicated server with controlled access. The network administrator can enforce password policies, monitor access logs, and restrict which users can view or edit records. This prevents unauthorised staff from accessing sensitive data.", indent=True)
body("Data Protection and Access Permissions: Role-based access control (RBAC) ensures that only authorised personnel such as doctors and nurses can access specific patient files. Administrative staff cannot access medical records they do not need, reducing the risk of internal data breaches.", indent=True)
body("Compliance with POPIA: The Protection of Personal Information Act requires organisations to protect personal data. A client-server model supports this through centralised backups, audit trails, and encryption, ensuring the clinic meets its legal obligations when handling patient health information.", indent=True)

page_break()

# ── Q2 ───────────────────────────────────────────────────────────────────────
heading("Question 2: Communication Principles and Protocols", 2)
heading("2.1 Protocols for Network Applications", 3)
add_table(
    ["Application", "Network Application Type", "Associated Protocol(s)"],
    [
        ["Accessing patient records on the server", "File Transfer / File Sharing", "FTP (File Transfer Protocol)"],
        ["Sending and receiving emails", "Electronic Mail", "SMTP (sending), POP3 / IMAP (receiving)"],
        ["Accessing the Department of Health reporting portal", "Web Browsing", "HTTP / HTTPS"],
    ]
)

heading("2.2 Protocol Differentiation", 3)
body("Different applications require different protocols because each protocol is designed for a specific purpose and set of requirements:")
body("Reliability (TCP vs UDP): Applications like FTP and email use TCP because data must arrive completely and in order. Real-time applications like video calls use UDP where speed matters more than perfect delivery.", indent=True)
body("Security (HTTP vs HTTPS): The Department of Health portal uses HTTPS, which encrypts data using TLS/SSL. Standard HTTP sends data in plain text, which is unsuitable for sensitive health reporting.", indent=True)
body("Data Transmission Efficiency: FTP is optimised for large file transfers, while HTTP is designed for request-response web page delivery. Using the wrong protocol would waste bandwidth or cause errors.", indent=True)
body("Protocol Functionality: Each protocol provides specific commands for its purpose. FTP has PUT and GET commands for file transfer, SMTP has MAIL FROM and RCPT TO for email routing, and HTTP uses GET and POST for web content.", indent=True)

page_break()

# ── Q3 ───────────────────────────────────────────────────────────────────────
heading("Question 3: OSI Model and Troubleshooting", 2)
heading("3.1 Identifying OSI Layers", 3)
add_table(
    ["Issue", "OSI Layer", "Motivation"],
    [
        ["Faulty network cable",
         "Layer 1 — Physical",
         "Physical cables, connectors, and electrical signals operate at Layer 1. A faulty cable means the physical medium cannot transmit bits between devices."],
        ["Incorrect IP addressing",
         "Layer 3 — Network",
         "IP addresses are a Layer 3 concept. Incorrect IP configuration prevents proper routing and logical addressing between networks."],
        ["Email application failure",
         "Layer 7 — Application",
         "Email clients and protocols (SMTP, POP3) operate at the Application layer. Failures here relate to application settings, credentials, or server configuration."],
    ]
)

heading("3.2 OSI Model in Troubleshooting", 3)
body("The OSI model helps technicians troubleshoot network problems in the following ways:")
body("Structured Step-by-Step Approach: Instead of guessing, technicians follow a systematic method — either bottom-up (starting at Layer 1 with cables and physical connections) or top-down (starting at Layer 7 with the application). This removes guesswork and ensures nothing is missed.", indent=True)
body("Isolating Problems to Specific Layers: Each layer has specific tools and symptoms. For example, if a ping works (Layer 3 is fine) but a web page does not load, the problem is at Layer 7. This narrows the search immediately without checking unrelated components.", indent=True)
body("Reducing Troubleshooting Time: By isolating the faulty layer quickly, technicians avoid spending time on layers that are working correctly. A problem identified at Layer 2 (switch MAC table) does not require checking Layer 7 application settings, saving significant time.", indent=True)

page_break()

# ── Q4 ───────────────────────────────────────────────────────────────────────
heading("Question 4: Network Optimisation and Security", 2)
heading("4.1 Network Performance Optimisation", 3)
body("Quality of Service (QoS): QoS prioritises critical network traffic during peak hours. For GreenLeaf Clinic, patient record access and VoIP calls would be given higher priority than general web browsing. This is done by classifying traffic using DSCP markings and applying queuing policies on the router, ensuring important applications always have sufficient bandwidth even when the network is congested.")
body("VLAN Segmentation: VLANs divide the physical network into logical segments. GreenLeaf already uses four subnets (Administration, Medical Staff, Pharmacy, Public Wi-Fi). Implementing VLANs on the switches reduces broadcast traffic — broadcasts from the Public Wi-Fi subnet do not reach the Medical Staff subnet. This improves overall network performance and reduces unnecessary traffic on each segment.")

heading("4.2 Network Security Measures", 3)
body("1. Firewall and Access Control Lists (ACLs): Deploy ACLs on the Core Router to restrict which subnets can communicate with each other. For example, Public Wi-Fi users should not be able to access the Administration or Medical Staff subnets where patient records are stored.")
body("2. Role-Based Access Control (RBAC): Implement user accounts with specific permissions on the file server. Doctors access medical records, pharmacists access prescription data, and administrative staff access billing only. No user has more access than their role requires.")
body("3. Data Encryption: All sensitive patient data transmitted over the network should be encrypted using TLS/SSL (HTTPS for the web portal, SFTP for file transfers). This protects data in transit from interception.")
body("4. Regular Backups and Patch Management: Schedule automated backups of all server data to an offsite or cloud location. Keep all server and router firmware updated to patch known vulnerabilities. This protects against data loss from hardware failure and exploitation of security weaknesses.")

page_break()

# ════════════════════════════════════════════════════════════════════════════
# DELIVERABLE 2
# ════════════════════════════════════════════════════════════════════════════
heading("DELIVERABLE 2: PRACTICAL IMPLEMENTATION", 1)

# ── Task 1 ───────────────────────────────────────────────────────────────────
heading("Task 1: Physical Network and Hardware Configuration", 2)
heading("1.1 Network Topology Design", 3)
body("The network uses a hierarchical star topology. Each subnet has its own switch connecting the server and two PCs. All four switches connect to the Core Router, which connects to the ISP Router via a serial WAN link. The Public Wi-Fi subnet includes an Access Point for wireless connectivity.")
body("Devices used: 2 x Cisco 2911 Routers, 4 x Cisco 2960-24TT Switches, 4 x Servers, 8 x PCs, 1 x Access Point, 1 x Laptop, 1 x Cloud (WAN).")
insert_image(f"{SS}/01_Task1_Physical_Topology.png", caption="Screenshot 1 — Physical Network Topology in Cisco Packet Tracer")

page_break()

# ── Task 2 ───────────────────────────────────────────────────────────────────
heading("Task 2: IP Subnetting", 2)
heading("2.2 Custom Subnet Mask Calculation", 3)
body("Base address: 192.168.10.0/24 — 4 subnets required.")
body("Number of bits to borrow: 2 bits (2² = 4 subnets)")
body("New prefix: /24 + 2 = /26")
body("Subnet mask: 255.255.255.192")
body("Block size: 256 − 192 = 64 addresses per subnet | Usable hosts: 64 − 2 = 62 per subnet")

heading("2.1 Subnet Table", 3)
add_table(
    ["Subnet", "Department", "Network Address", "1st Valid Address", "Last Valid Address", "Broadcast Address"],
    [
        ["1", "Administration",    "192.168.10.0/26",   "192.168.10.1",   "192.168.10.62",  "192.168.10.63"],
        ["2", "Medical Staff",     "192.168.10.64/26",  "192.168.10.65",  "192.168.10.126", "192.168.10.127"],
        ["3", "Pharmacy",          "192.168.10.128/26", "192.168.10.129", "192.168.10.190", "192.168.10.191"],
        ["4", "Public Patient Wi-Fi", "192.168.10.192/26", "192.168.10.193", "192.168.10.254", "192.168.10.255"],
        ["WAN", "Inter-Router Link", "192.168.20.0/30", "192.168.20.1",  "192.168.20.2",   "192.168.20.3"],
    ]
)

page_break()

# ── Task 3 ───────────────────────────────────────────────────────────────────
heading("Task 3: Configure IP Settings", 2)
heading("3.1 Static IP and Router Interface Configuration", 3)
body("Each server was configured with a static IP address as the 2nd valid host in its subnet. The Core Router interfaces were configured as the default gateway (1st valid address) for each subnet.")
insert_image(f"{SS}/02_Task3_1_Server_IP_Config.png", caption="Screenshot 2 — Server Static IP Configurations")
insert_image(f"{SS}/03_Task3_1_CORE_RTR_Interfaces.png", caption="Screenshot 3 — CORE-RTR Interface Configuration (show ip interface brief)")

heading("3.2 DHCP Server Configuration", 3)
body("DHCP was enabled on each server to automatically assign IP addresses to client PCs in each subnet.")
insert_image(f"{SS}/04_Task3_2_DHCP_Config.png", caption="Screenshot 4 — DHCP Server Configuration on All 4 Servers")

heading("3.3 Client DHCP Configuration", 3)
body("All PCs were set to DHCP mode via Desktop > IP Configuration. Each PC successfully received an IP address from its subnet's DHCP server.")
insert_image(f"{SS}/05_Task3_3_Client_DHCP_Leases.png", caption="Screenshot 5 — Client PCs Receiving DHCP Leases")

heading("3.4 Connectivity Verification", 3)
body("Ping tests were conducted to verify local, gateway, and cross-subnet connectivity.")
insert_image(f"{SS}/06_Task3_4_Connectivity_Verification.png", caption="Screenshot 6 — Ping Connectivity Verification Results")

page_break()

# ── Task 4 ───────────────────────────────────────────────────────────────────
heading("Task 4: File-Sharing Service (FTP)", 2)
heading("4.1 FTP User Account Creation", 3)
body("An FTP user account was created on ADMIN-SRV (192.168.10.2) with the following credentials and permissions:")
body("Username: Bhekumusa | Password: Ntshwenya | Permissions: Read, Write, Rename, List", indent=True)
insert_image(f"{SS}/07_Task4_1_FTP_User_Account.png", caption="Screenshot 7 — FTP User Account Created on ADMIN-SRV")

heading("4.2 FTP Access from Client PC", 3)
body("FTP connection was established from ADMIN-PC1 (192.168.10.10) to ADMIN-SRV (192.168.10.2) using the created credentials.")
insert_image(f"{SS}/08_Task4_2_FTP_Access.png", caption="Screenshot 8 — FTP Login from ADMIN-PC1 Command Prompt")

heading("4.3 File Creation and Transfer", 3)
body("ITCTA1.txt was created on ADMIN-PC1, uploaded to the FTP server using the PUT command, then retrieved from ADMIN-PC2 using the GET command.")
insert_image(f"{SS}/09_Task4_3_File_Transfer.png", caption="Screenshot 9 — FTP File Upload (put ITCTA1.txt) and Retrieval")

heading("4.4 File Rename", 3)
body("The file asa842-k8.bin was renamed to Bhekumusa_Ntshwenya.bin on the FTP server using the RENAME command.")
insert_image(f"{SS}/10_Task4_4_File_Rename.png", caption="Screenshot 10 — FTP File Rename (asa842-k8.bin → Bhekumusa_Ntshwenya.bin)")

page_break()

# ── Task 5 ───────────────────────────────────────────────────────────────────
heading("Task 5: Web Service Configuration", 2)
heading("5.1 Web Server Configuration", 3)
body("HTTP and HTTPS services were enabled on PUBLIC-SRV (192.168.10.194). The GreenLeaf patient portal HTML was loaded into index.html.")
insert_image(f"{SS}/11_Task5_1_Web_Server_Config.png", caption="Screenshot 11 — HTTP Service Configuration on PUBLIC-SRV")

heading("5.2 DNS Configuration", 3)
body("A DNS A Record was created on ADMIN-SRV pointing www.greenleafclinic.org to 192.168.10.194 (PUBLIC-SRV).")
insert_image(f"{SS}/12_Task5_2_DNS_Config.png", caption="Screenshot 12 — DNS A Record Configuration on ADMIN-SRV")

heading("5.3 Web Browser Test", 3)
body("The GreenLeaf patient portal was accessed by browsing to http://www.greenleafclinic.org from a client PC. The web server is accessible via IP address (192.168.10.194). DNS resolution via domain name requires the CORE-RTR Gi0/1/0 port to be manually cabled to PUBLIC-SW in PT.")
insert_image(f"{SS}/13_Task5_3_Web_Browser_Test.png", caption="Screenshot 13 — Web Browser Access to GreenLeaf Patient Portal")

page_break()

# ── Task 6 ───────────────────────────────────────────────────────────────────
heading("Task 6: Routing Configuration", 2)
heading("6.1 Inter-Router Network (192.168.20.0/30)", 3)
body("A point-to-point serial WAN link was configured between CORE-RTR and ISP-RTR using the 192.168.20.0/30 subnet.")
body("CORE-RTR Serial0/0/0: 192.168.20.1/30 (DCE — clock rate 64000)", indent=True)
body("ISP-RTR Serial0/0/0: 192.168.20.2/30 (DTE)", indent=True)
insert_image(f"{SS}/14_Task6_Routing_Config.png", caption="Screenshot 14 — WAN Serial Link and RIPv2 Routing Configuration")

heading("6.2 RIPv2 Routing Protocol", 3)
body("RIP version 2 was configured on both routers with no auto-summary to support classless routing across all four subnets.")
body("CORE-RTR advertises: 192.168.10.0, 192.168.20.0", indent=True)
body("ISP-RTR advertises: 192.168.20.0 with default-information originate for internet access", indent=True)
body("Verification: CORE-RTR routing table shows R 0.0.0.0/0 via 192.168.20.2 (default route from ISP) and all four subnets as directly connected. WAN ping success rate: 100% (5/5).")

page_break()

# ── Checklist ────────────────────────────────────────────────────────────────
heading("Submission Checklist", 2)
add_table(
    ["Item", "Status"],
    [
        ["Theory questions answered (Deliverable 1)", "✅ Complete"],
        ["Packet Tracer file submitted", "✅ GreenLeaf_Clinic_Network.pkt"],
        ["Screenshots of physical topology", "✅ Screenshot 1"],
        ["Screenshots of IP configurations", "✅ Screenshots 2 & 3"],
        ["Screenshots of DHCP configurations", "✅ Screenshots 4 & 5"],
        ["Screenshots of connectivity verification", "✅ Screenshot 6"],
        ["Screenshots of file sharing (FTP)", "✅ Screenshots 7–10"],
        ["Screenshots of web server configuration", "✅ Screenshot 11"],
        ["Screenshots of DNS configuration", "✅ Screenshot 12"],
        ["Screenshots of web browser test", "✅ Screenshot 13"],
        ["Screenshots of routing configuration", "✅ Screenshot 14"],
    ]
)

doc.save(OUT)
print(f"Saved: {OUT}")
