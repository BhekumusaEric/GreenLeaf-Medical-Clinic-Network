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
    style_run(run, bold=True, size={1:14,2:12,3:11}[level])

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
style_run(p.add_run("Deliverable 2 — Task 4: File-Sharing Service (FTP)"), size=12)

doc.add_paragraph()

add_table(
    ["Field", "Details"],
    [
        ["Student Name",   "Charlie Mandiwana"],
        ["Campus",         "Bedfordview"],
        ["Module",         "ITCTA1-44 — Network Architecture & Implementation"],
        ["Assessment",     "Cisco Packet Tracer Practical Assessment"],
        ["Date",           "November 2026"],
    ]
)

doc.add_page_break()

# ── TASK 4 ───────────────────────────────────────────────────────────────────
heading("Task 4: File-Sharing Service (FTP)", 1)

heading("4.1 FTP User Account Creation", 2)
body("A dedicated FTP user account was set up on the Administration Server (ADMIN-SRV) at IP address 192.168.10.2. The account was created with full file management permissions to allow secure access to patient records stored on the server.")
body("Username: Charlie  |  Password: Mandiwana  |  Permissions: Read, Write, Delete, Rename, List", indent=True)
insert_image(f"{SS}/07_Task4_1_FTP_User_Account.png", caption="Screenshot 7 — FTP User Account (Charlie Mandiwana) Created on ADMIN-SRV")

doc.add_page_break()

heading("4.2 FTP Access from Client PC", 2)
body("An FTP session was initiated from ADMIN-PC1 (192.168.10.10) to ADMIN-SRV (192.168.10.2) using the Command Prompt on the client PC. The connection was successfully established and the user Charlie logged in using the configured credentials.")
body("Command used: ftp 192.168.10.2", indent=True)
insert_image(f"{SS}/08_Task4_2_FTP_Access.png", caption="Screenshot 8 — Successful FTP Login from ADMIN-PC1 to ADMIN-SRV")

doc.add_page_break()

heading("4.3 File Creation and Transfer", 2)
body("A text file named ITCTA1.txt was created on ADMIN-PC1 and uploaded to the FTP server using the PUT command. The file was then retrieved from a second client machine (ADMIN-PC2) using the GET command, confirming that file sharing across the network was working correctly.")
body("Upload command:   put ITCTA1.txt", indent=True)
body("Download command: get ITCTA1.txt", indent=True)
insert_image(f"{SS}/09_Task4_3_File_Transfer.png", caption="Screenshot 9 — FTP File Upload and Retrieval (ITCTA1.txt)")

doc.add_page_break()

heading("4.4 File Rename on FTP Server", 2)
body("An existing file on the FTP server (asa842-k8.bin) was renamed to Charlie_Mandiwana.bin using the RENAME command. This demonstrates the ability to manage files remotely on the server through an FTP session, which is an important capability for maintaining organised patient record storage.")
body("Rename command: rename asa842-k8.bin Charlie_Mandiwana.bin", indent=True)
insert_image(f"{SS}/10_Task4_4_File_Rename.png", caption="Screenshot 10 — FTP File Rename (asa842-k8.bin → Charlie_Mandiwana.bin)")

doc.save(OUT)
print(f"Saved: {OUT}")
