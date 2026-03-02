import React, { useState, useMemo } from 'react';
import { Plus, Trash2, Cpu, HardDrive, Server, Info, Copy, CheckCircle2, Download } from 'lucide-react';

// Zoom Node official published VM specifications
const MODULES = [
  {
    id: 'znode_base',
    name: 'Zoom Node Management Server (Default)',
    cpu: 8,
    ram: 16,
    storage: 200,
    secStorage: 0,
    notes: 'Base requirement for deploying Zoom Node and standard modules.',
  },
  {
    id: 'znode_min',
    name: 'Zoom Node Server (Minimum/Testing)',
    cpu: 3,
    ram: 8,
    storage: 200,
    secStorage: 0,
    notes: 'Minimum viable configuration for testing or single-service lightweight deployments.',
  },
  {
    id: 'meeting_connector',
    name: 'Meeting Connector',
    cpu: 8,
    ram: 16,
    storage: 200,
    secStorage: 200,
    notes: 'Allows for ~400 concurrent attendees using Group HD (720p) for local meeting hosting.',
  },
  {
    id: 'meetings_hybrid',
    name: 'Meetings Hybrid',
    cpu: 8,
    ram: 16,
    storage: 200,
    secStorage: 200,
    notes: 'Allows for ~400 concurrent attendees using Group HD (720p) for hybrid routing.',
  },
  {
    id: 'zpls_2k',
    name: 'Phone Local Survivability (Up to 2,000 Reg)',
    cpu: 8,
    ram: 16,
    storage: 80,
    secStorage: 0,
    notes: 'Supports 240 concurrent calls, 2 CPS, 45 RPS.',
  },
  {
    id: 'zpls_5k',
    name: 'Phone Local Survivability (Up to 5,000 Reg)',
    cpu: 16,
    ram: 16,
    storage: 80,
    secStorage: 0,
    notes: 'Supports 480 concurrent calls, 4 CPS, 90 RPS.',
  },
  {
    id: 'rec_10',
    name: 'Recording Connector (10 Simultaneous)',
    cpu: 8,
    ram: 16,
    storage: 100,
    secStorage: 200,
    notes: 'Max 10 recordings, 2 transcodings. External NFS (500GB+) recommended.',
  },
  {
    id: 'rec_20',
    name: 'Recording Connector (20 Simultaneous)',
    cpu: 16,
    ram: 32,
    storage: 100,
    secStorage: 350,
    notes: 'Max 20 recordings, 4 transcodings. External NFS (500GB+) recommended.',
  },
  {
    id: 'rec_30',
    name: 'Recording Connector (30 Simultaneous)',
    cpu: 24,
    ram: 84,
    storage: 100,
    secStorage: 500,
    notes: 'Max 30 recordings, 6 transcodings. External NFS (500GB+) recommended.',
  },
  {
    id: 'meeting_survivability',
    name: 'Meeting Survivability Module',
    cpu: 3,
    ram: 8,
    storage: 180,
    secStorage: 0,
    notes: 'Dedicated specifications for the Meeting Survivability component.',
  },
  {
    id: 'recording_hybrid',
    name: 'Zoom Recording Hybrid',
    cpu: 8,
    ram: 16,
    storage: 200,
    secStorage: 200,
    notes: 'Manages hybrid recording routing and local storage integration.',
  },
  {
    id: 'team_chat_hybrid',
    name: 'Zoom Team Chat Hybrid',
    cpu: 8,
    ram: 16,
    storage: 200,
    secStorage: 0,
    notes: 'Provides local data loss prevention (DLP) and archiving for Team Chat.',
  },
  {
    id: 'transfer_service',
    name: 'Zoom Transfer Service',
    cpu: 4,
    ram: 8,
    storage: 100,
    secStorage: 0,
    notes: 'Facilitates secure file transfers and migrations within hybrid environments.',
  },
  {
    id: 'content_streamer',
    name: 'Zoom Content Streamer',
    cpu: 8,
    ram: 16,
    storage: 200,
    secStorage: 0,
    notes: 'Streamlines and optimizes local network content delivery for meetings/webinars.',
  }
];

const FIREWALL_RULES = {
  znode_base: [
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Zoom Cloud Communication' },
    { dir: 'Inbound', proto: 'TCP', port: '443, 8443', dest: 'Node IP', desc: 'Web Administration' },
    { dir: 'Inbound', proto: 'TCP', port: '22', dest: 'Node IP', desc: 'SSH / CLI Access' },
  ],
  znode_min: [
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Zoom Cloud Communication' },
    { dir: 'Inbound', proto: 'TCP', port: '443', dest: 'Node IP', desc: 'Web Administration' },
  ],
  meeting_connector: [
    { dir: 'Inbound', proto: 'TCP', port: '8801, 8802', dest: 'Module IP', desc: 'Meeting Client Control' },
    { dir: 'Inbound', proto: 'UDP', port: '8801, 8802', dest: 'Module IP', desc: 'Meeting Media (Audio/Video/Share)' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
  ],
  meetings_hybrid: [
    { dir: 'Inbound', proto: 'TCP', port: '8801, 8802', dest: 'Module IP', desc: 'Meeting Client Control' },
    { dir: 'Inbound', proto: 'UDP', port: '8801, 8802', dest: 'Module IP', desc: 'Meeting Media (Audio/Video/Share)' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
  ],
  zpls_2k: [
    { dir: 'Inbound', proto: 'TCP/UDP', port: '5060, 5061', dest: 'Module IP', desc: 'SIP Signaling' },
    { dir: 'Inbound', proto: 'UDP', port: '20000-65535', dest: 'Module IP', desc: 'RTP Media' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
  ],
  zpls_5k: [
    { dir: 'Inbound', proto: 'TCP/UDP', port: '5060, 5061', dest: 'Module IP', desc: 'SIP Signaling' },
    { dir: 'Inbound', proto: 'UDP', port: '20000-65535', dest: 'Module IP', desc: 'RTP Media' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
  ],
  rec_10: [
    { dir: 'Inbound', proto: 'TCP', port: '443', dest: 'Module IP', desc: 'Recording Web Access' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
    { dir: 'Outbound', proto: 'TCP/UDP', port: '2049, 111', dest: 'NFS Server IP', desc: 'NFS Storage Mount' },
  ],
  rec_20: [
    { dir: 'Inbound', proto: 'TCP', port: '443', dest: 'Module IP', desc: 'Recording Web Access' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
    { dir: 'Outbound', proto: 'TCP/UDP', port: '2049, 111', dest: 'NFS Server IP', desc: 'NFS Storage Mount' },
  ],
  rec_30: [
    { dir: 'Inbound', proto: 'TCP', port: '443', dest: 'Module IP', desc: 'Recording Web Access' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
    { dir: 'Outbound', proto: 'TCP/UDP', port: '2049, 111', dest: 'NFS Server IP', desc: 'NFS Storage Mount' },
  ],
  meeting_survivability: [
    { dir: 'Inbound', proto: 'TCP', port: '8801, 8802', dest: 'Module IP', desc: 'Meeting Client Control' },
    { dir: 'Inbound', proto: 'UDP', port: '8801, 8802', dest: 'Module IP', desc: 'Meeting Media (Audio/Video/Share)' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
  ],
  recording_hybrid: [
    { dir: 'Inbound', proto: 'TCP', port: '443', dest: 'Module IP', desc: 'Recording Web/API Access' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync & Management' },
    { dir: 'Outbound', proto: 'TCP/UDP', port: '2049, 111', dest: 'NFS Server IP', desc: 'NFS Storage Mount (if applicable)' },
  ],
  team_chat_hybrid: [
    { dir: 'Inbound', proto: 'TCP', port: '443', dest: 'Module IP', desc: 'Chat Client Connectivity' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync & Key Management' },
  ],
  transfer_service: [
    { dir: 'Inbound', proto: 'TCP', port: '443', dest: 'Module IP', desc: 'Secure File Transfer Access' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
  ],
  content_streamer: [
    { dir: 'Inbound', proto: 'TCP', port: '8801, 8802', dest: 'Module IP', desc: 'Content Client Control' },
    { dir: 'Inbound', proto: 'UDP', port: '8801, 8802', dest: 'Module IP', desc: 'Content Media Stream' },
    { dir: 'Outbound', proto: 'TCP', port: '443', dest: '*.zoom.us', desc: 'Cloud Sync' },
  ]
};

const DEPLOYMENT_GUIDES = {
  znode_base: [
    { step: 1, task: 'Provision Virtual Machine', detail: 'Deploy the Zoom Node OVF/VHD template to your hypervisor with the specified CPU, RAM, and Storage.' },
    { step: 2, task: 'Network Configuration', detail: 'Boot the VM and use the CLI to configure static IP, Subnet Mask, Gateway, and DNS servers.' },
    { step: 3, task: 'Web Console Access', detail: 'Navigate to https://<Node-IP>:8443 in your web browser and log in with default credentials.' },
    { step: 4, task: 'Generate Setup Token', detail: 'Log into the Zoom Web Portal, navigate to Node Management, and generate a new installation token.' },
    { step: 5, task: 'Register Node', detail: 'Paste the installation token into the local Zoom Node web console to register it to your Zoom account.' },
  ],
  znode_min: [
    { step: 1, task: 'Provision Virtual Machine', detail: 'Deploy the Zoom Node OVF/VHD template using the minimum specs (testing only).' },
    { step: 2, task: 'Network Configuration', detail: 'Boot the VM and use the CLI to configure static IP, Subnet Mask, Gateway, and DNS servers.' },
    { step: 3, task: 'Register Node', detail: 'Access the web console and register the node using a token from the Zoom Web Portal.' },
  ],
  meeting_connector: [
    { step: 1, task: 'Prerequisite Check', detail: 'Ensure the base Zoom Node VM is fully deployed, registered, and online in the Zoom Web Portal.' },
    { step: 2, task: 'Add Service Module', detail: 'In the Zoom Web Portal, navigate to Zoom Node > Services > Meeting Connector and click Add.' },
    { step: 3, task: 'Network Routing', detail: 'Configure external or internal routing depending on whether external participants need access.' },
    { step: 4, task: 'Install Service', detail: 'Push the module installation from the Web Portal. The Node will download and install the Meeting Connector container.' },
  ],
  meetings_hybrid: [
    { step: 1, task: 'Prerequisite Check', detail: 'Ensure the base Zoom Node VM is fully deployed, registered, and online in the Zoom Web Portal.' },
    { step: 2, task: 'Add Service Module', detail: 'In the Zoom Web Portal, navigate to Zoom Node > Services > Meetings Hybrid and click Add.' },
    { step: 3, task: 'Network Routing', detail: 'Configure external or internal routing depending on whether external participants need access.' },
    { step: 4, task: 'Install Service', detail: 'Push the module installation from the Web Portal. The Node will download and install the Meetings Hybrid container.' },
  ],
  zpls_2k: [
    { step: 1, task: 'Prerequisite Check', detail: 'Ensure the base Zoom Node VM is fully deployed, registered, and online.' },
    { step: 2, task: 'Add ZPLS Module', detail: 'In the Zoom Web Portal, deploy the Phone Local Survivability service to the registered Node.' },
    { step: 3, task: 'Configure SIP', detail: 'Configure SIP trunking details and survivability routing rules in the Zoom Phone admin portal.' },
    { step: 4, task: 'Assign to Site', detail: 'Assign this specific ZPLS node to the desired Zoom Phone site for failover.' },
  ],
  zpls_5k: [
    { step: 1, task: 'Prerequisite Check', detail: 'Ensure the base Zoom Node VM is fully deployed, registered, and online.' },
    { step: 2, task: 'Add ZPLS Module', detail: 'In the Zoom Web Portal, deploy the Phone Local Survivability service to the registered Node.' },
    { step: 3, task: 'Configure SIP', detail: 'Configure SIP trunking details and survivability routing rules in the Zoom Phone admin portal.' },
    { step: 4, task: 'Assign to Site', detail: 'Assign this specific ZPLS node to the desired Zoom Phone site for failover.' },
  ],
  rec_10: [
    { step: 1, task: 'Storage Preparation', detail: 'Deploy and configure an external NFS server (NFSv3 or NFSv4) with at least 500GB capacity.' },
    { step: 2, task: 'Deploy Module', detail: 'Add the Recording Connector service to your Node via the Zoom Web Portal.' },
    { step: 3, task: 'Mount Storage', detail: 'Configure the NFS mount point within the Recording Connector settings in the Web Portal.' },
    { step: 4, task: 'Routing Rules', detail: 'Configure Recording Connector routing rules to determine which meetings are captured locally.' },
  ],
  rec_20: [
    { step: 1, task: 'Storage Preparation', detail: 'Deploy and configure an external NFS server (NFSv3 or NFSv4) with adequate capacity.' },
    { step: 2, task: 'Deploy Module', detail: 'Add the Recording Connector service to your Node via the Zoom Web Portal.' },
    { step: 3, task: 'Mount Storage', detail: 'Configure the NFS mount point within the Recording Connector settings in the Web Portal.' },
    { step: 4, task: 'Routing Rules', detail: 'Configure Recording Connector routing rules to determine which meetings are captured locally.' },
  ],
  rec_30: [
    { step: 1, task: 'Storage Preparation', detail: 'Deploy and configure an external NFS server (NFSv3 or NFSv4) with adequate capacity.' },
    { step: 2, task: 'Deploy Module', detail: 'Add the Recording Connector service to your Node via the Zoom Web Portal.' },
    { step: 3, task: 'Mount Storage', detail: 'Configure the NFS mount point within the Recording Connector settings in the Web Portal.' },
    { step: 4, task: 'Routing Rules', detail: 'Configure Recording Connector routing rules to determine which meetings are captured locally.' },
  ],
  meeting_survivability: [
    { step: 1, task: 'Prerequisite Check', detail: 'Ensure the base Zoom Node VM is deployed and online.' },
    { step: 2, task: 'Deploy Module', detail: 'Install the Meeting Survivability module via the Zoom Web Portal.' },
    { step: 3, task: 'DNS Configuration', detail: 'Update local DNS records to point the survivability URLs to the Node IP address.' },
    { step: 4, task: 'Testing', detail: 'Disconnect the WAN link and verify clients failover to the local survivability module.' },
  ],
  recording_hybrid: [
    { step: 1, task: 'Prerequisite Check', detail: 'Ensure the base Zoom Node VM is fully deployed, registered, and online.' },
    { step: 2, task: 'Storage Preparation', detail: 'Deploy and configure an external NFS server if local offloading is required.' },
    { step: 3, task: 'Add Service Module', detail: 'In the Zoom Web Portal, deploy the Recording Hybrid service.' },
    { step: 4, task: 'Configure Policies', detail: 'Define recording routing rules in the Zoom Web Portal.' },
  ],
  team_chat_hybrid: [
    { step: 1, task: 'Prerequisite Check', detail: 'Ensure the base Zoom Node VM is fully deployed, registered, and online.' },
    { step: 2, task: 'Add Service Module', detail: 'In the Zoom Web Portal, deploy the Team Chat Hybrid service.' },
    { step: 3, task: 'Certificate Config', detail: 'Configure required TLS certificates for secure local client connections.' },
    { step: 4, task: 'Enable Hybrid Chat', detail: 'Toggle Team Chat Hybrid routing rules for the target users/groups in the portal.' },
  ],
  transfer_service: [
    { step: 1, task: 'Prerequisite Check', detail: 'Ensure the base Zoom Node VM is fully deployed, registered, and online.' },
    { step: 2, task: 'Add Service Module', detail: 'In the Zoom Web Portal, deploy the Transfer Service module.' },
    { step: 3, task: 'Network Routing', detail: 'Ensure DNS and firewalls allow appropriate internal file transfer routing.' },
  ],
  content_streamer: [
    { step: 1, task: 'Prerequisite Check', detail: 'Ensure the base Zoom Node VM is fully deployed, registered, and online.' },
    { step: 2, task: 'Add Service Module', detail: 'In the Zoom Web Portal, deploy the Content Streamer module.' },
    { step: 3, task: 'Network Routing', detail: 'Configure network routing to direct applicable media traffic to the streamer IP.' },
  ]
};

export default function App() {
  const [items, setItems] = useState([
    { id: crypto.randomUUID(), moduleId: 'znode_base', quantity: 1 }
  ]);
  const [copied, setCopied] = useState(false);

  const addItem = () => {
    setItems([...items, { id: crypto.randomUUID(), moduleId: 'znode_base', quantity: 1 }]);
  };

  const removeItem = (id) => {
    setItems(items.filter(item => item.id !== id));
  };

  const updateItem = (id, field, value) => {
    setItems(items.map(item => {
      if (item.id === id) {
        return { ...item, [field]: value };
      }
      return item;
    }));
  };

  const totals = useMemo(() => {
    return items.reduce(
      (acc, item) => {
        const moduleDef = MODULES.find(m => m.id === item.moduleId);
        if (moduleDef && item.quantity > 0) {
          acc.cpu += moduleDef.cpu * item.quantity;
          acc.ram += moduleDef.ram * item.quantity;
          acc.storage += moduleDef.storage * item.quantity;
          acc.secStorage += moduleDef.secStorage * item.quantity;
          
          // Only count base node instances as actual VMs
          if (item.moduleId === 'znode_base' || item.moduleId === 'znode_min') {
            acc.vms += parseInt(item.quantity);
          }
        }
        return acc;
      },
      { cpu: 0, ram: 0, storage: 0, secStorage: 0, vms: 0 }
    );
  }, [items]);

  const copyToClipboard = () => {
    const text = `Zoom Node VM Resource Requirements
----------------------------------
Total VMs: ${totals.vms}
Total vCPUs: ${totals.cpu} Cores
Total RAM: ${totals.ram} GB
Total Primary Storage: ${totals.storage} GB
Total Secondary Storage: ${totals.secStorage > 0 ? totals.secStorage + ' GB' : 'N/A'}

Line Items:
${items.map(item => {
  const mod = MODULES.find(m => m.id === item.moduleId);
  return `- ${item.quantity}x ${mod?.name} (${mod?.cpu} CPU, ${mod?.ram}GB RAM, ${mod?.storage}GB Storage/ea)`;
}).join('\n')}
`;
    
    // Workaround for restricted iframe environments
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.style.position = "fixed"; // Avoid scrolling
    textArea.style.left = "-9999px";
    textArea.style.top = "-9999px";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
      document.execCommand('copy');
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error('Fallback: Oops, unable to copy', err);
    }
    
    document.body.removeChild(textArea);
  };

  const exportCSV = () => {
    const headers = ["Module Name", "Quantity", "CPU/ea", "RAM/ea (GB)", "Primary Storage/ea (GB)", "Secondary Storage/ea (GB)"];
    const rows = items.map(item => {
      const mod = MODULES.find(m => m.id === item.moduleId);
      return [
        `"${mod?.name || 'Unknown'}"`,
        item.quantity,
        mod?.cpu || 0,
        mod?.ram || 0,
        mod?.storage || 0,
        mod?.secStorage || 0
      ];
    });
    const totalsRow = ["TOTALS", totals.vms, totals.cpu, totals.ram, totals.storage, totals.secStorage];
    
    const csvContent = [headers.join(","), ...rows.map(r => r.join(",")), totalsRow.join(",")].join("\n");
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", "zoom_node_specs.csv");
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const exportXLS = () => {
    const selectedModuleIds = [...new Set(items.map(i => i.moduleId))];

    // Helper to escape XML special characters
    const escapeXml = (unsafe) => {
      return (unsafe || "").toString().replace(/[<>&'"]/g, function (c) {
        switch (c) {
          case '<': return '&lt;';
          case '>': return '&gt;';
          case '&': return '&amp;';
          case '\'': return '&apos;';
          case '"': return '&quot;';
        }
      });
    };

    const buildCell = (val, type="String") => `<Cell><Data ss:Type="${type}">${escapeXml(val)}</Data></Cell>`;

    let xml = `<?xml version="1.0"?>
    <?mso-application progid="Excel.Sheet"?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel"
     xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
     xmlns:html="http://www.w3.org/TR/REC-html40">
    `;

    // Sheet 1: Main VM Specs
    xml += `<Worksheet ss:Name="VM Specs"><Table>`;
    xml += `<Row>${buildCell("Module Name")}${buildCell("Quantity")}${buildCell("CPU/ea")}${buildCell("RAM/ea (GB)")}${buildCell("Primary Storage/ea (GB)")}${buildCell("Secondary Storage/ea (GB)")}</Row>`;

    items.forEach(item => {
      const mod = MODULES.find(m => m.id === item.moduleId);
      xml += `<Row>${buildCell(mod?.name)}${buildCell(item.quantity, "Number")}${buildCell(mod?.cpu, "Number")}${buildCell(mod?.ram, "Number")}${buildCell(mod?.storage, "Number")}${buildCell(mod?.secStorage, "Number")}</Row>`;
    });

    xml += `<Row>${buildCell("TOTALS")}${buildCell(totals.vms, "Number")}${buildCell(totals.cpu, "Number")}${buildCell(totals.ram, "Number")}${buildCell(totals.storage, "Number")}${buildCell(totals.secStorage, "Number")}</Row>`;
    xml += `</Table></Worksheet>`;

    // Sheets 2..N: Firewall Rules and Deployment Guides per selected module
    selectedModuleIds.forEach(modId => {
      const mod = MODULES.find(m => m.id === modId);
      const rules = FIREWALL_RULES[modId] || [];
      const guides = DEPLOYMENT_GUIDES[modId] || [];
      
      if (mod) {
        if (rules.length > 0) {
          // Excel worksheet names must be <= 31 chars and cannot contain certain special characters
          let sheetName = ("FW-" + mod.name).substring(0, 31).replace(/[\\/?*[\]:]/g, ' ').trim();

          xml += `<Worksheet ss:Name="${escapeXml(sheetName)}"><Table>`;
          xml += `<Row>${buildCell("Direction")}${buildCell("Protocol")}${buildCell("Port(s)")}${buildCell("Destination")}${buildCell("Description")}</Row>`;

          rules.forEach(rule => {
            xml += `<Row>${buildCell(rule.dir)}${buildCell(rule.proto)}${buildCell(rule.port)}${buildCell(rule.dest)}${buildCell(rule.desc)}</Row>`;
          });
          
          xml += `</Table></Worksheet>`;
        }

        if (guides.length > 0) {
          let sheetName = ("Guide-" + mod.name).substring(0, 31).replace(/[\\/?*[\]:]/g, ' ').trim();

          xml += `<Worksheet ss:Name="${escapeXml(sheetName)}"><Table>`;
          xml += `<Row>${buildCell("Step")}${buildCell("Task")}${buildCell("Detail")}</Row>`;

          guides.forEach(guide => {
            xml += `<Row>${buildCell(guide.step, "Number")}${buildCell(guide.task)}${buildCell(guide.detail)}</Row>`;
          });
          
          xml += `</Table></Worksheet>`;
        }
      }
    });

    xml += `</Workbook>`;

    const blob = new Blob([xml], { type: 'application/vnd.ms-excel' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", "zoom_node_specs.xls");
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans text-slate-800">
      <div className="max-w-6xl mx-auto space-y-6">
        
        {/* Header */}
        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
          <div className="flex items-center gap-3 mb-2">
            <div className="p-3 bg-blue-100 text-blue-600 rounded-xl">
              <Server className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-2xl font-bold text-slate-900">Zoom Node VM Calculator</h1>
              <p className="text-sm text-slate-500">Calculate resource requirements based on published Zoom Node specifications</p>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          
          {/* Main List */}
          <div className="lg:col-span-2 space-y-4">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
              <div className="px-6 py-4 border-b border-slate-100 bg-slate-50 flex justify-between items-center">
                <h2 className="font-semibold text-slate-700">Virtual Machine Line Items</h2>
                <button
                  onClick={addItem}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg transition-colors"
                >
                  <Plus className="w-4 h-4" />
                  Add Module
                </button>
              </div>
              
              <div className="p-6 space-y-4">
                {items.length === 0 ? (
                  <div className="text-center py-8 text-slate-400">
                    No modules added. Click "Add Module" to start.
                  </div>
                ) : (
                  items.map((item, index) => {
                    const activeModule = MODULES.find(m => m.id === item.moduleId);
                    return (
                      <div key={item.id} className="relative p-4 rounded-xl border border-slate-200 bg-slate-50 hover:border-blue-300 transition-colors">
                        <div className="flex flex-col md:flex-row gap-4 items-start md:items-center">
                          
                          {/* Module Selector */}
                          <div className="flex-1 w-full">
                            <label className="block text-xs font-medium text-slate-500 mb-1 uppercase tracking-wider">Service Module</label>
                            <select
                              value={item.moduleId}
                              onChange={(e) => updateItem(item.id, 'moduleId', e.target.value)}
                              className="w-full bg-white border border-slate-300 text-slate-700 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                            >
                              {MODULES.map(mod => (
                                <option key={mod.id} value={mod.id}>{mod.name}</option>
                              ))}
                            </select>
                          </div>

                          {/* Quantity */}
                          <div className="w-full md:w-24">
                            <label className="block text-xs font-medium text-slate-500 mb-1 uppercase tracking-wider">Quantity</label>
                            <input
                              type="number"
                              min="1"
                              value={item.quantity}
                              onChange={(e) => updateItem(item.id, 'quantity', parseInt(e.target.value) || 0)}
                              className="w-full bg-white border border-slate-300 text-slate-700 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                            />
                          </div>

                          {/* Delete Action */}
                          <div className="pt-5 hidden md:block">
                            <button
                              onClick={() => removeItem(item.id)}
                              className="p-2 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-colors"
                              title="Remove Item"
                            >
                              <Trash2 className="w-5 h-5" />
                            </button>
                          </div>
                        </div>

                        {/* Mobile Delete */}
                        <button
                          onClick={() => removeItem(item.id)}
                          className="absolute top-4 right-4 md:hidden text-slate-400 hover:text-red-500"
                        >
                          <Trash2 className="w-5 h-5" />
                        </button>

                        {/* Spec Hint */}
                        {activeModule && (
                          <div className="mt-3 text-xs text-slate-500 flex items-start gap-1">
                            <Info className="w-4 h-4 text-blue-400 shrink-0" />
                            <span>
                              <strong>Specs per VM:</strong> {activeModule.cpu} Cores, {activeModule.ram}GB RAM, {activeModule.storage}GB Storage {activeModule.secStorage > 0 && `(+${activeModule.secStorage}GB Sec)`}
                              <br className="hidden md:block"/>
                              <span className="md:ml-1 text-slate-400">{activeModule.notes}</span>
                            </span>
                          </div>
                        )}
                      </div>
                    );
                  })
                )}
              </div>
            </div>
          </div>

          {/* Sidebar / Totals */}
          <div className="lg:col-span-1">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 sticky top-8">
              <div className="p-6 border-b border-slate-100">
                <h2 className="text-lg font-bold text-slate-800">Total Requirements</h2>
                <p className="text-sm text-slate-500">Combined infrastructure needed</p>
              </div>
              
              <div className="p-6 space-y-6">
                
                {/* VM Count */}
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3 text-slate-600">
                    <Server className="w-5 h-5" />
                    <span className="font-medium">Total VMs</span>
                  </div>
                  <span className="text-xl font-bold text-slate-900">{totals.vms}</span>
                </div>

                {/* CPU */}
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3 text-slate-600">
                    <Cpu className="w-5 h-5" />
                    <span className="font-medium">Total vCPUs</span>
                  </div>
                  <span className="text-xl font-bold text-slate-900">{totals.cpu} <span className="text-sm font-normal text-slate-500">Cores</span></span>
                </div>

                {/* RAM */}
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3 text-slate-600">
                    <div className="w-5 h-5 flex items-center justify-center font-bold text-xs border-2 border-current rounded-sm">R</div>
                    <span className="font-medium">Total Memory</span>
                  </div>
                  <span className="text-xl font-bold text-slate-900">{totals.ram} <span className="text-sm font-normal text-slate-500">GB</span></span>
                </div>

                {/* Storage */}
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3 text-slate-600">
                    <HardDrive className="w-5 h-5" />
                    <span className="font-medium">Primary Storage</span>
                  </div>
                  <span className="text-xl font-bold text-slate-900">{totals.storage} <span className="text-sm font-normal text-slate-500">GB</span></span>
                </div>

                {/* Sec Storage (Conditional) */}
                {totals.secStorage > 0 && (
                  <div className="flex items-center justify-between pt-4 border-t border-slate-100">
                    <div className="flex items-center gap-3 text-slate-600">
                      <HardDrive className="w-5 h-5 text-slate-400" />
                      <span className="font-medium">Secondary Volume</span>
                    </div>
                    <span className="text-xl font-bold text-slate-900">{totals.secStorage} <span className="text-sm font-normal text-slate-500">GB</span></span>
                  </div>
                )}
                
                <button
                  onClick={copyToClipboard}
                  className="w-full mt-4 flex items-center justify-center gap-2 px-4 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium rounded-xl transition-colors"
                >
                  {copied ? <CheckCircle2 className="w-5 h-5 text-green-600" /> : <Copy className="w-5 h-5" />}
                  {copied ? 'Copied to Clipboard' : 'Copy Summary'}
                </button>
                
                <div className="grid grid-cols-2 gap-3 mt-3">
                  <button
                    onClick={exportCSV}
                    className="w-full flex items-center justify-center gap-2 px-4 py-3 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-xl transition-colors text-sm"
                  >
                    <Download className="w-4 h-4" />
                    CSV
                  </button>
                  <button
                    onClick={exportXLS}
                    className="w-full flex items-center justify-center gap-2 px-4 py-3 bg-blue-600 hover:bg-blue-700 text-white font-medium rounded-xl transition-colors text-sm"
                  >
                    <Download className="w-4 h-4" />
                    XLS
                  </button>
                </div>
                
              </div>
              <div className="px-6 py-4 bg-slate-50 text-xs text-slate-400 rounded-b-2xl border-t border-slate-100">
                Note: Ensure network interfaces support at least 10 Gbps and host hardware matches Zoom's recommended processor generation (e.g. Intel 3rd gen Xeon+).
              </div>
            </div>
          </div>

        </div>
      </div>
    </div>
  );
}
