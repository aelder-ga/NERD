export interface Classification { domain: string; fn: string; description: string; terms: string[]; }
export const CLASSIFICATIONS: Classification[] = [
  {
    "domain": "NETWORK",
    "fn": "Physical Networking",
    "description": "Cabling, ports, physical connectivity",
    "terms": [
      "cable",
      "cabling",
      "fiber",
      "copper",
      "patch panel",
      "patch cord",
      "wall jack",
      "switch port",
      "uplink"
    ]
  },
  {
    "domain": "NETWORK",
    "fn": "Logical Networking",
    "description": "VLANs, routing, logical topology",
    "terms": [
      "VLAN",
      "subnet",
      "routing",
      "gateway",
      "trunk",
      "tagging",
      "network topology",
      "segmentation"
    ]
  },
  {
    "domain": "NETWORK",
    "fn": "Wireless",
    "description": "Wireless infrastructure and connectivity",
    "terms": [
      "Wi-Fi",
      "wireless",
      "SSID",
      "access point",
      "AP",
      "roaming",
      "coverage",
      "signal",
      "wireless testing"
    ]
  },
  {
    "domain": "NETWORK",
    "fn": "Firewall",
    "description": "Firewall configuration and traffic control",
    "terms": [
      "firewall",
      "ACL",
      "traffic rule",
      "NAT",
      "port forwarding",
      "network access rule"
    ]
  },
  {
    "domain": "NETWORK",
    "fn": "Internet",
    "description": "External internet connectivity and services",
    "terms": [
      "internet",
      "ISP",
      "circuit",
      "broadband",
      "WAN provider",
      "external connectivity"
    ]
  },
  {
    "domain": "NETWORK",
    "fn": "Voice",
    "description": "Telephone and voice communication systems",
    "terms": [
      "telephone",
      "phone",
      "Webex calling",
      "PSTN",
      "hunt group",
      "voicemail",
      "extension",
      "auto attendant",
      "SIP"
    ]
  },
  {
    "domain": "SERVER",
    "fn": "Hardware",
    "description": "Physical server hardware and systems",
    "terms": [
      "server hardware",
      "rack server",
      "CPU",
      "memory",
      "RAID controller",
      "replacement server"
    ]
  },
  {
    "domain": "SERVER",
    "fn": "Virtualization",
    "description": "Virtual servers and hosting infrastructure",
    "terms": [
      "VM",
      "virtual machine",
      "hypervisor",
      "Hyper-V",
      "virtual host",
      "guest",
      "virtualization"
    ]
  },
  {
    "domain": "SERVER",
    "fn": "Storage",
    "description": "Server and infrastructure storage systems",
    "terms": [
      "SAN",
      "NAS",
      "storage array",
      "volume",
      "disk",
      "datastore",
      "capacity"
    ]
  },
  {
    "domain": "SERVER",
    "fn": "Network Services",
    "description": "DNS and DHCP network services",
    "terms": [
      "DNS",
      "DHCP",
      "resolver",
      "name resolution",
      "lease",
      "scope",
      "reservation"
    ]
  },
  {
    "domain": "ENDPOINT",
    "fn": "Desktop",
    "description": "Fixed workstation computing devices",
    "terms": [
      "desktop",
      "workstation",
      "MiniPC",
      "fixed PC"
    ]
  },
  {
    "domain": "ENDPOINT",
    "fn": "Laptop",
    "description": "Portable conventional computing devices",
    "terms": [
      "laptop",
      "notebook",
      "portable PC",
      "docking station"
    ]
  },
  {
    "domain": "ENDPOINT",
    "fn": "Chromebook",
    "description": "ChromeOS student and staff devices",
    "terms": [
      "Chromebook",
      "ChromeOS",
      "student device",
      "enrollment",
      "Chromebook repair"
    ]
  },
  {
    "domain": "ENDPOINT",
    "fn": "iPad",
    "description": "District-managed Apple tablet devices",
    "terms": [
      "iPad",
      "iPadOS",
      "Apple tablet"
    ]
  },
  {
    "domain": "ENDPOINT",
    "fn": "Cellular",
    "description": "Cellular-connected phones and hotspots",
    "terms": [
      "cell phone",
      "cellular",
      "mobile phone",
      "SIM",
      "hotspot",
      "mobile data"
    ]
  },
  {
    "domain": "ENDPOINT",
    "fn": "Presentation",
    "description": "Interactive display and presentation systems",
    "terms": [
      "interactive display",
      "smart board",
      "projector",
      "classroom display",
      "screen casting"
    ]
  },
  {
    "domain": "ENDPOINT",
    "fn": "Printing",
    "description": "Printing, scanning, and MFP systems",
    "terms": [
      "printer",
      "printing",
      "MFP",
      "scanner",
      "scanning",
      "print queue",
      "print driver"
    ]
  },
  {
    "domain": "ENDPOINT",
    "fn": "Peripheral",
    "description": "Attached and supporting endpoint devices",
    "terms": [
      "keyboard",
      "mouse",
      "webcam",
      "headset",
      "adapter",
      "peripheral",
      "USB device"
    ]
  },
  {
    "domain": "IDENTITY",
    "fn": "Account Management",
    "description": "Account creation and lifecycle management",
    "terms": [
      "account creation",
      "account disable",
      "separation",
      "account deletion",
      "staff offboarding",
      "account lifecycle"
    ]
  },
  {
    "domain": "IDENTITY",
    "fn": "Provisioning",
    "description": "Account synchronization between systems",
    "terms": [
      "account sync",
      "provisioning",
      "directory sync",
      "automated account creation",
      "account mapping"
    ]
  },
  {
    "domain": "IDENTITY",
    "fn": "Authentication",
    "description": "Verifying user identity",
    "terms": [
      "sign-in",
      "login",
      "MFA",
      "Duo",
      "password",
      "SSO",
      "identity verification",
      "authentication"
    ]
  },
  {
    "domain": "IDENTITY",
    "fn": "Standard Authorization",
    "description": "Normal user access controls",
    "terms": [
      "permissions",
      "group membership",
      "user access",
      "sharing access",
      "application access",
      "authorization"
    ]
  },
  {
    "domain": "IDENTITY",
    "fn": "Administrative Authorization",
    "description": "Elevated user access controls",
    "terms": [
      "admin role",
      "elevated privilege",
      "privileged access",
      "local administrator",
      "global administrator"
    ]
  },
  {
    "domain": "DATA",
    "fn": "Entry",
    "description": "Manual data creation and maintenance",
    "terms": [
      "data entry",
      "manual update",
      "form entry",
      "student data maintenance"
    ]
  },
  {
    "domain": "DATA",
    "fn": "Rostering",
    "description": "System-to-system data exchange",
    "terms": [
      "rostering",
      "roster sync",
      "Clever",
      "SIS export",
      "student import",
      "data transfer"
    ]
  },
  {
    "domain": "DATA",
    "fn": "Process Flow",
    "description": "Movement, handoffs, responsible roles",
    "terms": [
      "handoff",
      "data flow",
      "workflow",
      "process owner",
      "responsibility",
      "data movement"
    ]
  },
  {
    "domain": "DATA",
    "fn": "Validation",
    "description": "Verifying data accuracy and completeness",
    "terms": [
      "data validation",
      "accuracy",
      "completeness",
      "reconciliation",
      "duplicate data",
      "quality check"
    ]
  },
  {
    "domain": "DATA",
    "fn": "Data Governance",
    "description": "Defining data ownership and authority",
    "terms": [
      "data owner",
      "authoritative source",
      "data authority",
      "data stewardship",
      "data governance"
    ]
  },
  {
    "domain": "DATA",
    "fn": "State Reporting",
    "description": "Required state data submissions",
    "terms": [
      "state reporting",
      "NJSLEDS",
      "state submission",
      "required state data"
    ]
  },
  {
    "domain": "DATA",
    "fn": "Local Reporting",
    "description": "District-generated information and reports",
    "terms": [
      "district report",
      "local report",
      "dashboard data",
      "internal analysis",
      "district statistics"
    ]
  },
  {
    "domain": "CYBERSECURITY",
    "fn": "Prevention",
    "description": "Measures preventing cybersecurity incidents",
    "terms": [
      "hardening",
      "security baseline",
      "attack prevention",
      "protective control",
      "anti-malware"
    ]
  },
  {
    "domain": "CYBERSECURITY",
    "fn": "Detection",
    "description": "Monitoring, logging, and identification",
    "terms": [
      "security monitoring",
      "alert",
      "threat detection",
      "log review",
      "suspicious activity"
    ]
  },
  {
    "domain": "CYBERSECURITY",
    "fn": "Patch Management",
    "description": "Security, firmware, and software updates",
    "terms": [
      "patch",
      "security update",
      "firmware update",
      "software update",
      "vulnerability remediation"
    ]
  },
  {
    "domain": "CYBERSECURITY",
    "fn": "Backup",
    "description": "Preserving recoverable copies of data",
    "terms": [
      "backup",
      "backup schedule",
      "retention",
      "recoverable copy",
      "backup verification"
    ]
  },
  {
    "domain": "CYBERSECURITY",
    "fn": "Restoration",
    "description": "Restoring systems to operational readiness",
    "terms": [
      "disaster recovery",
      "ransomware recovery",
      "server rebuild",
      "VM restoration",
      "service restoration",
      "system rebuild",
      "configuration restoration",
      "bare-metal recovery",
      "replacement server",
      "restore from backup",
      "recovery testing"
    ]
  },
  {
    "domain": "CYBERSECURITY",
    "fn": "Incident Response",
    "description": "Responding to cybersecurity incidents",
    "terms": [
      "security incident",
      "ransomware response",
      "containment",
      "forensics",
      "incident triage",
      "breach response"
    ]
  },
  {
    "domain": "CYBERSECURITY",
    "fn": "Awareness",
    "description": "Educating users about security risks",
    "terms": [
      "security awareness",
      "phishing training",
      "safe computing",
      "user education"
    ]
  },
  {
    "domain": "CYBERSECURITY",
    "fn": "Audit & Controls",
    "description": "Verifying required protections are effective",
    "terms": [
      "security audit",
      "access audit",
      "disabled account audit",
      "control validation",
      "compliance check"
    ]
  },
  {
    "domain": "FACILITIES",
    "fn": "Access Control",
    "description": "Physical entry and access systems",
    "terms": [
      "door access",
      "badge",
      "card reader",
      "physical access",
      "door controller"
    ]
  },
  {
    "domain": "FACILITIES",
    "fn": "Surveillance",
    "description": "Cameras and video monitoring systems",
    "terms": [
      "camera",
      "CCTV",
      "surveillance",
      "video monitoring",
      "camera recording"
    ]
  },
  {
    "domain": "FACILITIES",
    "fn": "Visitor Management",
    "description": "Visitor screening and identification systems",
    "terms": [
      "visitor",
      "visitor screening",
      "Raptor",
      "guest badge",
      "visitor check-in"
    ]
  },
  {
    "domain": "FACILITIES",
    "fn": "Building Controls",
    "description": "Building automation and control systems",
    "terms": [
      "building automation",
      "HVAC control",
      "building controller"
    ]
  },
  {
    "domain": "FACILITIES",
    "fn": "Environmental Monitoring",
    "description": "Power and environmental condition monitoring",
    "terms": [
      "UPS",
      "power monitoring",
      "temperature sensor",
      "humidity",
      "leak sensor",
      "environmental alert"
    ]
  },
  {
    "domain": "FACILITIES",
    "fn": "Signage",
    "description": "Digital and electronic signage systems",
    "terms": [
      "digital signage",
      "electronic sign",
      "information display"
    ]
  },
  {
    "domain": "FACILITIES",
    "fn": "Public Address",
    "description": "Paging, speaker, and announcement systems",
    "terms": [
      "paging",
      "PA",
      "loudspeaker",
      "building announcement",
      "intercom"
    ]
  },
  {
    "domain": "FACILITIES",
    "fn": "Clocks",
    "description": "Building clock and synchronization systems",
    "terms": [
      "clock",
      "bell clock",
      "clock synchronization",
      "time display"
    ]
  },
  {
    "domain": "FACILITIES",
    "fn": "Timekeeping",
    "description": "Employee time tracking systems",
    "terms": [
      "time clock",
      "employee time tracking",
      "punch clock",
      "timesheet system"
    ]
  },
  {
    "domain": "BUSINESS",
    "fn": "Purchasing",
    "description": "Technology purchasing and procurement processes",
    "terms": [
      "procurement",
      "purchasing",
      "purchase order",
      "quote",
      "requisition"
    ]
  },
  {
    "domain": "BUSINESS",
    "fn": "Budgeting",
    "description": "Technology budgeting and financial planning",
    "terms": [
      "budget",
      "financial planning",
      "forecast",
      "technology spending"
    ]
  },
  {
    "domain": "BUSINESS",
    "fn": "Contracts",
    "description": "Agreements, renewals, and contract management",
    "terms": [
      "contract",
      "agreement",
      "renewal",
      "vendor terms"
    ]
  },
  {
    "domain": "BUSINESS",
    "fn": "Licensing",
    "description": "Product entitlement and usage requirements",
    "terms": [
      "license",
      "entitlement",
      "subscription rights",
      "activation allowance",
      "license compliance"
    ]
  },
  {
    "domain": "BUSINESS",
    "fn": "Asset Management",
    "description": "Inventory, lifecycle, and disposal management",
    "terms": [
      "inventory",
      "asset tag",
      "device lifecycle",
      "disposal",
      "surplus",
      "asset register"
    ]
  },
  {
    "domain": "BUSINESS",
    "fn": "Point of Sale",
    "description": "Payment and transaction processing systems",
    "terms": [
      "payment",
      "transaction",
      "POS",
      "card processing",
      "register payment"
    ]
  },
  {
    "domain": "BUSINESS",
    "fn": "Personnel",
    "description": "Technology staff administrative requirements",
    "terms": [
      "staff administration",
      "personnel process",
      "staffing requirement",
      "employee administrative record"
    ]
  },
  {
    "domain": "PUBLIC",
    "fn": "Parent Support",
    "description": "Direct parent assistance and communication",
    "terms": [
      "parent assistance",
      "parent help",
      "family support",
      "parent instructions"
    ]
  },
  {
    "domain": "PUBLIC",
    "fn": "Notifications",
    "description": "Targeted and mass audience messaging",
    "terms": [
      "mass notification",
      "targeted message",
      "Thrillshare",
      "broadcast notice",
      "audience message"
    ]
  },
  {
    "domain": "PUBLIC",
    "fn": "Emergency Communications",
    "description": "Urgent district status and alerts",
    "terms": [
      "emergency alert",
      "school closure notice",
      "urgent district status",
      "emergency message"
    ]
  },
  {
    "domain": "PUBLIC",
    "fn": "Website",
    "description": "Public web content and services",
    "terms": [
      "website",
      "public web page",
      "web publishing",
      "district site"
    ]
  },
  {
    "domain": "PUBLIC",
    "fn": "Portal",
    "description": "Public and parent-facing online services",
    "terms": [
      "parent portal",
      "public portal",
      "parent-facing service"
    ]
  },
  {
    "domain": "PUBLIC",
    "fn": "Calendar",
    "description": "District and school event publishing",
    "terms": [
      "calendar",
      "school event",
      "district event",
      "event publishing"
    ]
  },
  {
    "domain": "PUBLIC",
    "fn": "Social Media",
    "description": "District social media account management",
    "terms": [
      "social media",
      "district social account",
      "social post",
      "account publishing"
    ]
  },
  {
    "domain": "PUBLIC",
    "fn": "Streaming",
    "description": "Live and recorded public broadcasts",
    "terms": [
      "livestream",
      "live broadcast",
      "recorded broadcast",
      "board meeting stream"
    ]
  }
];
const tokens = (v: string): string[] => v.toLowerCase().replace(/accounts\b/g,'account').replace(/permissions\b/g,'permission').replace(/[^a-z0-9]+/g,' ').trim().split(/\s+/).filter(t=>t.length>1);
export function classify(query: string): (Classification & {matches: string[]; score: number})[] {
 const words = tokens(query).filter(t=>!['the','for','how','and','with','genesis','entra','microsoft','meraki'].includes(t));
 if (!words.length) return [];
 return CLASSIFICATIONS.map(c=> {
 const matches=c.terms.filter(term=>tokens(term).some(t=>words.includes(t)));
 const hay=tokens(c.domain+' '+c.fn+' '+c.description+' '+c.terms.join(' '));
 const score=words.reduce((n,w)=>n+(hay.includes(w)?1:0),0)+matches.filter(m=>query.toLowerCase().includes(m.toLowerCase())).length*0.25;
 return {...c,matches,score};
 }).filter(c=>c.score>0).sort((a,b)=>b.score-a.score).slice(0,6);
}

