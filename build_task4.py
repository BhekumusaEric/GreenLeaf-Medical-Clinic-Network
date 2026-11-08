from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

OUT = "/Users/eric/GreenLeaf-Medical-Clinic-Network/Charlie_Mandiwana_Task4.docx"
SS  = "/Users/eric/GreenLeaf-Medical-Clinic-Network/screenshots"
FONT = "Arial"

doc = Document()

for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1.18)
    section.right_margin  = Inches(1.18)

def style_run(run, bold=False, size=11, color=None):
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
    style_run(run, bold=True, size={1:14, 2:12, 3:11}[level])

def body(text, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(4)
    if indent:
        p.paragraph_format.left_indent = Inches(0.3)
    run = p.add_run(text)
    style_run(run)

def insert_image(path, caption=None):
    if os.path.exists(path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(path, width=Inches(6.0))
        if caption:
            cp = doc.add_paragraph(caption)
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in cp.runs:
                style_run(r, size=9)
    else:
        body(f"[Image not found: {path}]")

def add_table(headers, rows):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = "Table Grid"
    for i, h in enumerate(t.rows[0].cells):
        h.text = headers[i]
        for para in h.paragraphs:
            for run in para.runs:
                style_run(run, bold=True, size=10)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(t.rows[ri+1].cells):
            val.text = row[ci]
            for para in val.paragraphs:
                for run in para.runs:
                    style_run(run, size=10)
    doc.add_paragraph()

# ── COVER PAGE ───────────────────────────────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
style_run(p.add_run("Eduvos"), bold=True, size=13)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
style_run(p.add_run("Network Architecture & Implementation (ITCTA1-44)"), size=11)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
style_run(p.add_run("GREENLEAF MEDICAL CLINIC NETWORK"), bold=True, size=16)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
style_run(p.add_run("Deliverable 2 — Practical Implementation"), size=12)

doc.add_paragraph()

add_table(
    ["Field", "Details"],
    [
        ["Student Name",  "Charlie Mandiwana"],
        ["Campus",        "Bedfordview"],
        ["Module",        "ITCTA1-44 — Network Architecture & Implementation"],
        ["Assessment",    "Cisco Packet Tracer Practical Assessment"],
        ["Date",          "November 2026"],
    ]
)

doc.add_page_break()

# ════════════════════════════════════════════════════════════════════════════
heading("DELIVERABLE 2: PRACTICAL IMPLEMENTATION", 1)

# ── TASK 1 ───────────────────────────────────────────────────────────────────
heading("Task 1: Physical Network and Hardware Configuration", 2)
heading("1.1 Network Topology Design", 3)
body("For this task I had to build the full network topology in Cisco Packet Tracer. I used a star topology where each department has its own switch that connects back to the Core Router. The Core Router then connects to the ISP Router through a serial WAN link which simulates an internet connection.")
body("The devices I used were:")
body("2 x Cisco 2911 Routers (CORE-RTR and ISP-RTR)", indent=True)
body("4 x Cisco 2960-24TT Switches (one per department)", indent=True)
body("4 x Servers (ADMIN-SRV, MED-SRV, PHARM-SRV, PUBLIC-SRV)", indent=True)
body("8 x PCs (2 per department)", indent=True)
body("1 x Access Point (AP1) for the Public Wi-Fi subnet", indent=True)
body("1 x Laptop (LT1) connected wirelessly", indent=True)
body("1 x Cloud device (WAN) to represent the internet", indent=True)
insert_image(f"{SS}/01_Task1_Physical_Topology.png", caption="Screenshot 1 — Full Network Topology in Cisco Packet Tracer")

doc.add_page_break()

# ── TASK 2 ───────────────────────────────────────────────────────────────────
heading("Task 2: IP Subnetting", 2)
heading("2.1 Subnet Mask Calculation", 3)
body("The clinic was given the base network address 192.168.10.0/24 and I needed to divide it into 4 subnets — one for each department. To do this I had to borrow 2 bits from the host portion of the address.")
body("Bits borrowed: 2 (because 2² = 4 subnets)", indent=True)
body("New subnet mask: /26 which is 255.255.255.192", indent=True)
body("Block size: 256 - 192 = 64 addresses per subnet", indent=True)
body("Usable hosts per subnet: 64 - 2 = 62", indent=True)
body("I also configured a /30 WAN link between the two routers using the 192.168.20.0 network which gives only 2 usable host addresses — one for each router end.")

heading("2.2 Subnet Table", 3)
add_table(
    ["Subnet", "Department", "Network Address", "Default Gateway", "Server IP", "Broadcast"],
    [
        ["1", "Administration",      "192.168.10.0/26",   "192.168.10.1",   "192.168.10.2",   "192.168.10.63"],
        ["2", "Medical Staff",       "192.168.10.64/26",  "192.168.10.65",  "192.168.10.66",  "192.168.10.127"],
        ["3", "Pharmacy",            "192.168.10.128/26", "192.168.10.129", "192.168.10.130", "192.168.10.191"],
        ["4", "Public Patient Wi-Fi","192.168.10.192/26", "192.168.10.193", "192.168.10.194", "192.168.10.255"],
        ["WAN", "Inter-Router Link", "192.168.20.0/30",   "192.168.20.1",   "192.168.20.2",   "192.168.20.3"],
    ]
)

doc.add_page_break()

# ── TASK 3 ───────────────────────────────────────────────────────────────────
heading("Task 3: Configure IP Settings", 2)

heading("3.1 Static IP and Router Interface Configuration", 3)
body("Each server was given a static IP address. I used the second valid host address in each subnet for the server and the first valid host address as the default gateway on the router interface. This is standard practice so that the gateway IP is always predictable.")
body("I configured all four interfaces on CORE-RTR using the CLI with the ip address command and then used no shutdown to bring each interface up. The screenshot below shows all interfaces are up and have the correct IP addresses assigned.")
insert_image(f"{SS}/02_Task3_1_Server_IP_Config.png", caption="Screenshot 2 — Server Static IP Configuration (ADMIN-SRV shown)")
insert_image(f"{SS}/03_Task3_1_CORE_RTR_Interfaces.png", caption="Screenshot 3 — CORE-RTR show ip interface brief confirming all interfaces are up")

heading("3.2 DHCP Server Configuration", 3)
body("Instead of manually setting IP addresses on every PC, I configured each server to act as a DHCP server for its subnet. This means when a PC starts up it automatically gets an IP address, subnet mask, default gateway and DNS server from the server. I set up the DHCP pool on each server through the Services tab in Packet Tracer.")
insert_image(f"{SS}/04_Task3_2_DHCP_Config.png", caption="Screenshot 4 — DHCP Pool Configuration on ADMIN-SRV")

heading("3.3 Client DHCP Lease Verification", 3)
body("After setting the PCs to DHCP mode I checked that they received IP addresses correctly. I opened the IP Configuration on ADMIN-PC1 and confirmed it got an address in the 192.168.10.0/26 range from the DHCP server.")
insert_image(f"{SS}/05_Task3_3_Client_DHCP_Leases.png", caption="Screenshot 5 — ADMIN-PC1 showing DHCP lease received successfully")

heading("3.4 Connectivity Verification", 3)
body("To make sure the network was working I ran a ping test from ADMIN-PC1 to the default gateway (192.168.10.1). All 4 packets were received with 0% packet loss which confirmed that the PC could communicate with the router successfully.")
insert_image(f"{SS}/06_Task3_4_Connectivity_Verification.png", caption="Screenshot 6 — Ping test from ADMIN-PC1 to gateway 192.168.10.1 showing 100% success")

doc.add_page_break()

# ── TASK 4 ───────────────────────────────────────────────────────────────────
heading("Task 4: File-Sharing Service (FTP)", 2)

heading("4.1 FTP User Account Creation", 3)
body("I set up an FTP service on ADMIN-SRV so that staff could securely transfer patient files across the network. I created a user account with full permissions so the user can read, write, delete, rename and list files on the server.")
body("Username: Charlie  |  Password: Mandiwana  |  Permissions: Read, Write, Delete, Rename, List", indent=True)
insert_image(f"{SS}/07_Task4_1_FTP_User_Account.png", caption="Screenshot 7 — FTP User Account Created on ADMIN-SRV")

heading("4.2 FTP Login from Client PC", 3)
body("I tested the FTP service by connecting from ADMIN-PC1 to ADMIN-SRV using the Command Prompt. I typed ftp 192.168.10.2 and logged in with the Charlie account. The connection was successful and I was able to access the server.")
insert_image(f"{SS}/08_Task4_2_FTP_Access.png", caption="Screenshot 8 — Successful FTP Login from ADMIN-PC1 Command Prompt")

heading("4.3 File Transfer", 3)
body("I created a file called ITCTA1.txt on the client PC and uploaded it to the FTP server using the put command. I then used the dir command to confirm the file was on the server. This shows that the file sharing service is working correctly across the network.")
insert_image(f"{SS}/09_Task4_3_File_Transfer.png", caption="Screenshot 9 — FTP File Upload using put command")

heading("4.4 File Rename", 3)
body("I also tested the rename function by renaming an existing file on the server. I used the rename command to change asa842-k8.bin to Charlie_Mandiwana.bin. This confirms that the FTP permissions are set correctly and the user has full control over files on the server.")
insert_image(f"{SS}/10_Task4_4_File_Rename.png", caption="Screenshot 10 — FTP File Rename (asa842-k8.bin renamed to Charlie_Mandiwana.bin)")

doc.add_page_break()

# ── TASK 5 ───────────────────────────────────────────────────────────────────
heading("Task 5: Web Service Configuration", 2)

heading("5.1 Web Server Setup", 3)
body("I configured PUBLIC-SRV as the clinic's web server by enabling the HTTP and HTTPS services in the Services tab. I also loaded the GreenLeaf patient portal HTML page into the index.html file on the server so that patients can access the clinic's website.")
insert_image(f"{SS}/11_Task5_1_Web_Server_Config.png", caption="Screenshot 11 — HTTP and HTTPS Services Enabled on PUBLIC-SRV")

heading("5.2 DNS Configuration", 3)
body("To allow users to access the website using a domain name instead of an IP address, I set up a DNS record on ADMIN-SRV. I created an A record that maps www.greenleafclinic.org to the web server's IP address 192.168.10.194.")
insert_image(f"{SS}/12_Task5_2_DNS_Config.png", caption="Screenshot 12 — DNS A Record: www.greenleafclinic.org → 192.168.10.194")

heading("5.3 Web Browser Test", 3)
body("I tested the web server by opening the web browser on PUBLIC-PC1 and navigating to http://192.168.10.194. The GreenLeaf Medical Clinic patient portal loaded successfully, confirming that the web service is working correctly.")
insert_image(f"{SS}/13_Task5_3_Web_Browser_Test.png", caption="Screenshot 13 — GreenLeaf Patient Portal Loaded in Web Browser")

doc.add_page_break()

# ── TASK 6 ───────────────────────────────────────────────────────────────────
heading("Task 6: Routing Configuration", 2)

heading("6.1 WAN Serial Link", 3)
body("To connect the two routers I configured a point-to-point serial WAN link using the 192.168.20.0/30 subnet. CORE-RTR was set as the DCE end with a clock rate of 64000 and ISP-RTR was the DTE end. This link simulates the connection between the clinic and its internet service provider.")
body("CORE-RTR Serial0/0/0: 192.168.20.1/30 — DCE (clock rate 64000)", indent=True)
body("ISP-RTR Serial0/0/0: 192.168.20.2/30 — DTE", indent=True)

heading("6.2 RIPv2 Routing", 3)
body("I configured RIP version 2 on both routers so they could share routing information with each other. I used no auto-summary to make sure the classless subnet information was advertised correctly. On ISP-RTR I also added default-information originate so that a default route to the internet would be sent to CORE-RTR automatically.")
body("After the configuration I ran show ip route on CORE-RTR to verify the routing table. It showed the default route (0.0.0.0/0) received from ISP-RTR via RIP and all four internal subnets as directly connected routes.")
insert_image(f"{SS}/14_Task6_Routing_Config.png", caption="Screenshot 14 — CORE-RTR Routing Table showing RIPv2 routes and default route")

doc.save(OUT)
print(f"Saved: {OUT}")
