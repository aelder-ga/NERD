# TLDR: Search-Only Classification Reference

> DRAFT • Prepared 2026-09-30 • Based on DORK Framework v0.2 • No permanent ID assigned.

Search this page for the words you know (Ctrl+F / Find on page), then evaluate the suggested Domain and Function. **Why does this document exist?** governs the decision. Terms below are proposed discovery aids, not additional taxonomy, mandatory tags or automatic classification rules. A system name alone does not determine a document’s classification.

| Domain | Function | Common search terms |
| --- | --- | --- |
| NETWORK | Physical Networking | cable, cabling, fiber, copper, patch panel, patch cord, wall jack, switch port, uplink |
| NETWORK | Logical Networking | VLAN, subnet, routing, gateway, trunk, tagging, network topology, segmentation |
| NETWORK | Wireless | Wi-Fi, wireless, SSID, access point, AP, roaming, coverage, signal, wireless testing |
| NETWORK | Firewall | firewall, ACL, traffic rule, NAT, port forwarding, network access rule |
| NETWORK | Internet | internet, ISP, circuit, broadband, WAN provider, external connectivity |
| NETWORK | Voice | telephone, phone, Webex calling, PSTN, hunt group, voicemail, extension, auto attendant, SIP |
| SERVER | Hardware | server hardware, rack server, CPU, memory, RAID controller, replacement server |
| SERVER | Virtualization | VM, virtual machine, hypervisor, Hyper-V, virtual host, guest, virtualization |
| SERVER | Storage | SAN, NAS, storage array, volume, disk, datastore, capacity |
| SERVER | Network Services | DNS, DHCP, resolver, name resolution, lease, scope, reservation |
| ENDPOINT | Desktop | desktop, workstation, MiniPC, fixed PC |
| ENDPOINT | Laptop | laptop, notebook, portable PC, docking station |
| ENDPOINT | Chromebook | Chromebook, ChromeOS, student device, enrollment, Chromebook repair |
| ENDPOINT | iPad | iPad, iPadOS, Apple tablet |
| ENDPOINT | Cellular | cell phone, cellular, mobile phone, SIM, hotspot, mobile data |
| ENDPOINT | Presentation | interactive display, smart board, projector, classroom display, screen casting |
| ENDPOINT | Printing | printer, printing, MFP, scanner, scanning, print queue, print driver |
| ENDPOINT | Peripheral | keyboard, mouse, webcam, headset, adapter, peripheral, USB device |
| IDENTITY | Account Management | account creation, account disable, separation, account deletion, staff offboarding, account lifecycle |
| IDENTITY | Provisioning | account sync, provisioning, directory sync, automated account creation, account mapping |
| IDENTITY | Authentication | sign-in, login, MFA, Duo, password, SSO, identity verification, authentication |
| IDENTITY | Standard Authorization | permissions, group membership, user access, sharing access, application access, authorization |
| IDENTITY | Administrative Authorization | admin role, elevated privilege, privileged access, local administrator, global administrator |
| DATA | Entry | data entry, manual update, form entry, student data maintenance |
| DATA | Rostering | rostering, roster sync, Clever, SIS export, student import, data transfer |
| DATA | Process Flow | handoff, data flow, workflow, process owner, responsibility, data movement |
| DATA | Validation | data validation, accuracy, completeness, reconciliation, duplicate data, quality check |
| DATA | Data Governance | data owner, authoritative source, data authority, data stewardship, data governance |
| DATA | State Reporting | state reporting, NJSLEDS, state submission, required state data |
| DATA | Local Reporting | district report, local report, dashboard data, internal analysis, district statistics |
| CYBERSECURITY | Prevention | hardening, security baseline, attack prevention, protective control, anti-malware |
| CYBERSECURITY | Detection | security monitoring, alert, threat detection, log review, suspicious activity |
| CYBERSECURITY | Patch Management | patch, security update, firmware update, software update, vulnerability remediation |
| CYBERSECURITY | Backup | backup, backup schedule, retention, recoverable copy, backup verification |
| CYBERSECURITY | Restoration | disaster recovery, ransomware recovery, server rebuild, VM restoration, service restoration, system rebuild, configuration restoration, bare-metal recovery, replacement server, restore from backup, recovery testing |
| CYBERSECURITY | Incident Response | security incident, ransomware response, containment, forensics, incident triage, breach response |
| CYBERSECURITY | Awareness | security awareness, phishing training, safe computing, user education |
| CYBERSECURITY | Audit & Controls | security audit, access audit, disabled account audit, control validation, compliance check |
| FACILITIES | Access Control | door access, badge, card reader, physical access, door controller |
| FACILITIES | Surveillance | camera, CCTV, surveillance, video monitoring, camera recording |
| FACILITIES | Visitor Management | visitor, visitor screening, Raptor, guest badge, visitor check-in |
| FACILITIES | Building Controls | building automation, HVAC control, building controller |
| FACILITIES | Environmental Monitoring | UPS, power monitoring, temperature sensor, humidity, leak sensor, environmental alert |
| FACILITIES | Signage | digital signage, electronic sign, information display |
| FACILITIES | Public Address | paging, PA, loudspeaker, building announcement, intercom |
| FACILITIES | Clocks | clock, bell clock, clock synchronization, time display |
| FACILITIES | Timekeeping | time clock, employee time tracking, punch clock, timesheet system |
| BUSINESS | Purchasing | procurement, purchasing, purchase order, quote, requisition |
| BUSINESS | Budgeting | budget, financial planning, forecast, technology spending |
| BUSINESS | Contracts | contract, agreement, renewal, vendor terms |
| BUSINESS | Licensing | license, entitlement, subscription rights, activation allowance, license compliance |
| BUSINESS | Asset Management | inventory, asset tag, device lifecycle, disposal, surplus, asset register |
| BUSINESS | Point of Sale | payment, transaction, POS, card processing, register payment |
| BUSINESS | Personnel | staff administration, personnel process, staffing requirement, employee administrative record |
| PUBLIC | Parent Support | parent assistance, parent help, family support, parent instructions |
| PUBLIC | Notifications | mass notification, targeted message, Thrillshare, broadcast notice, audience message |
| PUBLIC | Emergency Communications | emergency alert, school closure notice, urgent district status, emergency message |
| PUBLIC | Website | website, public web page, web publishing, district site |
| PUBLIC | Portal | parent portal, public portal, parent-facing service |
| PUBLIC | Calendar | calendar, school event, district event, event publishing |
| PUBLIC | Social Media | social media, district social account, social post, account publishing |
| PUBLIC | Streaming | livestream, live broadcast, recorded broadcast, board meeting stream |

## When the same term appears in several places

| Topic | Purpose of this document | Classification to consider |
| --- | --- | --- |
| MFA / Duo | Explain how a person signs in | IDENTITY / Authentication |
| MFA / Duo | Verify required protections are effective | CYBERSECURITY / Audit & Controls |
| Employee separation | Disable or remove the account | IDENTITY / Account Management |
| Employee separation | Audit whether former staff still have access | CYBERSECURITY / Audit & Controls |
| Server rebuild | Explain a recovery operation | CYBERSECURITY / Restoration |
| Server hardware | Record hardware specifications | SERVER / Hardware |
| Standardized testing | Test wireless readiness | NETWORK / Wireless; Standardized Testing Collection |
| Standardized testing | Verify testing accommodations data | DATA / Validation; Standardized Testing Collection |
| Entra ID / Genesis / Meraki | Any of several purposes | Select Domain/Function by purpose; add applicable System / Platform separately |

Terms can overlap. For instance, “enrollment” may describe device enrollment, account provisioning or a student-data workflow. Read the document’s purpose before choosing.

Search terms are draft editorial additions except the Restoration example supplied in Framework v0.2. Revise them when real searches reveal gaps. They do not imply every listed product has a separate Function or every matching document belongs there.

See the [compact Classification Index](classification-index.md) and [usage guidance](usage-guide.md). This page helps locate classifications; it does not establish tenant search indexing or a completed dashboard.
