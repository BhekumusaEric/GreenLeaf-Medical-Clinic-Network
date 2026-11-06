# GreenLeaf Medical Clinic Network Upgrade
## Cisco Packet Tracer Assessment — Eduvos ITCTA1-44

This repository contains the complete network design, configuration scripts, subnetting calculations, and student answer sheet for the **GreenLeaf Medical Clinic Network** assessment.

---

## 📁 Repository Contents

* **[`GreenLeaf_Medical_Clinic_Student_Answer_Sheet.docx`](./GreenLeaf_Medical_Clinic_Student_Answer_Sheet.docx)**  
  The full, comprehensive student answer sheet covering Deliverable 1 (Networking Core Concepts, 50 Marks) and Deliverable 2 (Practical Implementation, 50 Marks) with 14 designated screenshot placeholder boxes.
* **[`CORE-RTR_Config.txt`](./CORE-RTR_Config.txt)**  
  Ready-to-paste Cisco IOS CLI configuration script for the central **Core Router** (interfaces, descriptions, WAN serial clock rate, and RIPv2 routing).
* **[`ISP-RTR_Config.txt`](./ISP-RTR_Config.txt)**  
  Ready-to-paste Cisco IOS CLI configuration script for the **ISP Gateway Router** (serial WAN connection and default route propagation).
* **[`Quick_Server_Setup.txt`](./Quick_Server_Setup.txt)**  
  Cheat sheet for configuring static IPs and GUI services across all 4 servers (DHCP, DNS, FTP, and Web HTTP).
* **[`greenleaf_index.html`](./greenleaf_index.html)**  
  Custom HTML website code for the clinic's patient portal, compliant with POPIA regulations.

---

## 🏥 Scenario Overview

GreenLeaf Medical Clinic is an expanding healthcare facility requiring a modern, secure, and resilient client-server network architecture:
* **4 Segregated Subnets:** Administration, Medical Staff, Pharmacy, and Public Patient Wi-Fi.
* **Core Router & ISP Gateway Router:** Connecting internal subnets and routing traffic to the Internet over a `/30` point-to-point serial link using RIPv2.
* **Dedicated Services:**
  * 4 × DHCP Servers (one per subnet).
  * 1 × Central DNS Server on Admin Server (`192.168.10.2`) resolving `www.greenleafclinic.org`.
  * 1 × Central File Server (FTP) for sensitive patient records on Admin Server.
  * 1 × Email Server (SMTP/POP3) on Pharmacy Server (`192.168.10.130`).
  * 1 × Web Server hosting the clinic portal on Public Patient Server (`192.168.10.194`).

---

## 🔢 IP Subnetting Scheme (`192.168.10.0/24`)

* **Custom Subnet Mask:** `255.255.255.192` (`/26`)
* **Block Size:** 64 | **Usable Hosts:** 62 per subnet

| Subnet # | Department | Network Address | Default Gateway | Server IP | Usable Host Range | Broadcast Address |
| :---: | :--- | :--- | :--- | :--- | :--- | :--- |
| **1** | Administration | `192.168.10.0/26` | `192.168.10.1` | `192.168.10.2` | `192.168.10.1` – `.62` | `192.168.10.63` |
| **2** | Medical Staff | `192.168.10.64/26` | `192.168.10.65` | `192.168.10.66` | `192.168.10.65` – `.126` | `192.168.10.127` |
| **3** | Pharmacy | `192.168.10.128/26` | `192.168.10.129` | `192.168.10.130` | `192.168.10.129` – `.190` | `192.168.10.191` |
| **4** | Public Patient Wi-Fi | `192.168.10.192/26` | `192.168.10.193` | `192.168.10.194` | `192.168.10.193` – `.254` | `192.168.10.255` |
| **WAN**| Inter-Router Link | `192.168.20.0/30` | `192.168.20.1` | `192.168.20.2` | `192.168.20.1` – `.2` | `192.168.20.3` |

---

## 🚀 Quick Setup Instructions

1. **Topology:** Create the network in Cisco Packet Tracer with 2 Routers, 4 Switches, 4 Servers, 8 PCs, 1 Access Point, and 1 Laptop.
2. **Core Router CLI:** Copy and paste [`CORE-RTR_Config.txt`](./CORE-RTR_Config.txt) into the CORE-RTR CLI.
3. **ISP Router CLI:** Copy and paste [`ISP-RTR_Config.txt`](./ISP-RTR_Config.txt) into the ISP-RTR CLI.
4. **Servers:** Follow [`Quick_Server_Setup.txt`](./Quick_Server_Setup.txt) to configure IP addresses and services.
5. **Web Portal:** Import [`greenleaf_index.html`](./greenleaf_index.html) into `index.html` on `PUBLIC-SRV`.
6. **Clients:** Set PCs to DHCP; connect the Laptop to the AP with wireless card `WPC300N` and set to DHCP.
7. **Verification & Screenshots:** Capture the 14 required screenshots and paste them into [`GreenLeaf_Medical_Clinic_Student_Answer_Sheet.docx`](./GreenLeaf_Medical_Clinic_Student_Answer_Sheet.docx).
8. **Save:** Save the Packet Tracer project as `GreenLeaf_Clinic_Network.pkt`.

---
*Created by Bhekumusa Eric Ntshwenya*
