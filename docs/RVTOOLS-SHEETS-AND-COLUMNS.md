# RVTools Sheets & Columns Reference

Every sheet and column from a real RVTools export. What Invar covers today is
tracked in `docs/PARITY-PLAN.md`, not here.

Use this to (a) name your sheets and columns exactly as RVTools does, and
(b) know upfront which columns are cheap and which are expensive.

**Column counts below exclude RVTools boilerplate** repeated on nearly every
sheet: `VI SDK Server`, `VI SDK UUID`, `VI SDK Server type`, `VI SDK API Version`,
`com.vrlcm.snapshot`, `Internal Sort Column`, `Object ID`. Also common across
most sheets and omitted from per-sheet notes: `Datacenter`, `Cluster`, `Host`,
`Folder`, `Annotation`, `VM ID`, `VM UUID`, `Powerstate`, `Template`,
`SRM Placeholder`, `OS according to the configuration file`,
`OS according to the VMware Tools`.

`Datacenter` and `Cluster` need walking the inventory tree from each object up
to its parent, which is a different query pattern than a flat property read.
Invar does that once per fetch, in the inventory path index.

---

## Sheet summary

| Sheet | RVTools cols | Notes |
|---|---:|---|
| vInfo | 89 | |
| vCPU | 30 | |
| vMemory | 34 | |
| vDisk | 40 | |
| vPartition | 22 | |
| vNetwork | 27 | |
| vCD | 21 | |
| vUSB | 26 | |
| vSnapshot | 22 | |
| vTools | 30 | |
| vSource | 12 | |
| vRP | 46 | |
| vCluster | 32 | |
| vHost | 70 | |
| vHBA | 11 | |
| vNIC | 12 | |
| vSwitch | 21 | |
| vPort | 20 | |
| dvSwitch | 27 | |
| dvPort | 38 | |
| vSC_VMK | 13 | |
| vDatastore | 27 | |
| vMultiPath | 32 | Needs HostMultipathInfo / storage path enumeration |
| vLicense | 8 | |
| vFileInfo | 5 | Needs HostDatastoreBrowser + SearchDatastoreSubFolders (datastore file-tree walking) |
| vHealth | 3 | Not alarms: RVTools' own computed checks; see below |
| vMetaData | 4 | |

---

## xlsx format parity

Read out of `reference/RVTools_export_all_2024-08-18_15.54.15.xlsx` (RVTools 4.6)
and reproduced by `src-tauri/src/export.rs`. Verified 2026-08-31 by generating an
export from the lab and reading it back with openpyxl.

| Element | RVTools | Ours |
|---|---|---|
| Header row | Verdana 9pt bold, `FFFFFFFF` on solid `FF000000`, left aligned | same |
| Body text | Verdana 9pt, general alignment | same |
| Integers | `#,##0`, right aligned | same |
| Decimals | `#,##0.00`, right aligned | same |
| Dates | real Excel serials, `yyyy/MM/dd HH:mm:ss` | same |
| Booleans | the text `True` / `False` — **not** Excel booleans | same |
| Freeze panes | `B2` (first row *and* first column) | same |
| AutoFilter | `A1` to the last cell, header included | same |
| Default row height | 15 | same |
| Sheet order | fixed RVTools order | same order, sheets we lack are skipped |
| Column widths | autofit to content | autofit to content (differs by data, as RVTools' would) |

RVTools also styles a column uniformly rather than per cell, so the format is
chosen from the whole column: a `Number` column is `#,##0` only if every value
is integral, and a `Text` column becomes a date column only if every non-empty
value parses as RFC 3339.

Two deliberate differences:

- **`vMetaData` keeps RVTools' four column names** (`RVTools major version`,
  `RVTools version`, `xlsx creation datetime`, `Server`) so RVTools-aware
  tooling still parses the file, but the values name *this* tool. Writing an
  RVTools version number would misstate where the data came from.
- The freeze pane is emitted as `state="frozen"` where RVTools writes
  `state="frozenSplit"`. Both are frozen panes at `B2` in Excel;
  `rust_xlsxwriter` does not expose the split variant.

Timestamps are written as vCenter reports them (UTC), not converted to local
time — converting would silently relabel every date in the export.

---

## Per-sheet detail

For each sheet: the full RVTools column list, plus anything about the sheet that
the list alone does not tell you.


### vInfo

**RVTools columns:**

```
VM, Powerstate, Template, SRM Placeholder, Config status, DNS Name, Connection state, Guest state, Heartbeat, Consolidation Needed, PowerOn, Suspended To Memory, Suspend time, Suspend Interval, Creation date, Change Version, CPUs, Overall Cpu Readiness, Memory, Active Memory, NICs, Disks, Total disk capacity MiB, Fixed Passthru HotPlug, min Required EVC Mode Key, Latency Sensitivity, Op Notification Timeout, EnableUUID, CBT, Primary IP Address, Network #1, Network #2, Network #3, Network #4, Network #5, Network #6, Network #7, Network #8, Num Monitors, Video Ram KiB, Resource pool, Folder ID, Folder, vApp, DAS protection, FT State, FT Role, FT Latency, FT Bandwidth, FT Sec. Latency, Vm Failover In Progress, Provisioned MiB, In Use MiB, Unshared MiB, HA Restart Priority, HA Isolation Response, HA VM Monitoring, Cluster rule(s), Cluster rule name(s), Boot Required, Boot delay, Boot retry delay, Boot retry enabled, Boot BIOS setup, Reboot PowerOff, EFI Secure boot, Firmware, HW version, HW upgrade status, HW upgrade policy, HW target, Path, Log directory, Snapshot directory, Suspend directory, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, OS according to the configuration file, OS according to the VMware Tools, Customization Info, Guest Detailed Data, VM ID, SMBIOS UUID, VM UUID
```


### vCPU

**RVTools columns:**

```
VM, Powerstate, Template, SRM Placeholder, CPUs, Sockets, Cores p/s, Max, Overall, Level, Shares, Reservation, Entitlement, DRS Entitlement, Limit, Hot Add, Hot Remove, Numa Hotadd Exposed, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, Folder, OS according to the configuration file, OS according to the VMware Tools, VM ID, VM UUID
```


### vMemory

**RVTools columns:**

```
VM, Powerstate, Template, SRM Placeholder, Size MiB, Memory Reservation Locked To Max, Overhead, Max, Consumed, Consumed Overhead, Private, Shared, Swapped, Ballooned, Active, Entitlement, DRS Entitlement, Level, Shares, Reservation, Limit, Hot Add, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, Folder, OS according to the configuration file, OS according to the VMware Tools, VM ID, VM UUID
```


### vDisk

**RVTools columns:**

```
VM, Powerstate, Template, SRM Placeholder, Disk, Disk Key, Disk UUID, Disk Path, Capacity MiB, Raw, Disk Mode, Sharing mode, Thin, Eagerly Scrub, Split, Write Through, Level, Shares, Reservation, Limit, Controller, Label, SCSI Unit #, Unit #, Shared Bus, Path, Raw LUN ID, Raw Comp. Mode, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, Folder, OS according to the configuration file, OS according to the VMware Tools, VM ID, VM UUID
```


### vPartition

**RVTools columns:**

```
VM, Powerstate, Template, SRM Placeholder, Disk Key, Disk, Capacity MiB, Consumed MiB, Free MiB, Free %, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, Folder, OS according to the configuration file, OS according to the VMware Tools, VM ID, VM UUID
```


### vNetwork

**RVTools columns:**

```
VM, Powerstate, Template, SRM Placeholder, NIC label, Adapter, Network, Switch, Connected, Starts Connected, Mac Address, Type, IPv4 Address, IPv6 Address, Direct Path IO, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, Folder, OS according to the configuration file, OS according to the VMware Tools, VM ID, VM UUID
```


### vCD

**RVTools columns:**

```
VM, Powerstate, Template, SRM Placeholder, Device Node, Connected, Starts Connected, Device Type, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, Folder, OS according to the configuration file, OS according to the VMware Tools, VMRef, VM ID, VM UUID
```


### vUSB

**RVTools columns:**

```
VM, Powerstate, Template, SRM Placeholder, Device Node, Device Type, Connected, Family, Speed, EHCI enabled, Auto connect, Bus number, Unit number, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, Folder, OS according to the configuration file, OS according to the VMware tools, VMRef, VM ID, VM UUID
```


### vSnapshot

**RVTools columns:**

```
VM, Powerstate, Name, Description, Date / time, Filename, Size MiB (vmsn), Size MiB (total), Quiesced, State, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, Folder, OS according to the configuration file, OS according to the VMware Tools, VM ID, VM UUID
```


### vTools

**RVTools columns:**

```
VM, Powerstate, Template, SRM Placeholder, VM Version, Tools, Tools Version, Required Version, Upgradeable, Upgrade Policy, Sync time, App status, Heartbeat status, Kernel Crash state, Operation Ready, State change support, Interactive Guest, Annotation, Application, LOCATION, Reboot, Datacenter, Cluster, Host, Folder, OS according to the configuration file, OS according to the VMware Tools, VMRef, VM ID, VM UUID
```


### vSource

**RVTools columns:**

```
Name, OS type, API type, API version, Version, Patch level, Build, Fullname, Product name, Product version, Product line, Vendor
```


### vRP

**RVTools columns:**

```
Resource Pool name, Resource Pool path, Status, # VMs total, # VMs, # vCPUs, CPU limit, CPU overheadLimit, CPU reservation, CPU level, CPU shares, CPU expandableReservation, CPU maxUsage, CPU overallUsage, CPU reservationUsed, CPU reservationUsedForVm, CPU unreservedForPool, CPU unreservedForVm, Mem Configured, Mem limit, Mem overheadLimit, Mem reservation, Mem level, Mem shares, Mem expandableReservation, Mem maxUsage, Mem overallUsage, Mem reservationUsed, Mem reservationUsedForVm, Mem unreservedForPool, Mem unreservedForVm, QS overallCpuDemand, QS overallCpuUsage, QS staticCpuEntitlement, QS distributedCpuEntitlement, QS balloonedMemory, QS compressedMemory, QS consumedOverheadMemory, QS distributedMemoryEntitlement, QS guestMemoryUsage, QS hostMemoryUsage, QS overheadMemory, QS privateMemory, QS sharedMemory, QS staticMemoryEntitlement, QS swappedMemory
```


### vCluster

**RVTools columns:**

```
Name, Config status, OverallStatus, NumHosts, numEffectiveHosts, TotalCpu, NumCpuCores, NumCpuThreads, Effective Cpu, TotalMemory, Effective Memory, Num VMotions, HA enabled, Failover Level, AdmissionControlEnabled, Host monitoring, HB Datastore Candidate Policy, Isolation Response, Restart Priority, Cluster Settings, Max Failures, Max Failure Window, Failure Interval, Min Up Time, VM Monitoring, DRS enabled, DRS default VM behavior, DRS vmotion rate, DPM enabled, DPM default behavior, DPM Host Power Action Rate, com.vmware.vcenter.cluster.edrs.upgradeHostAdded
```


### vHost

**RVTools columns:**

```
Host, Datacenter, Cluster, Config status, Compliance Check State, in Maintenance Mode, in Quarantine Mode, vSAN Fault Domain Name, CPU Model, Speed, HT Available, HT Active, # CPU, Cores per CPU, # Cores, CPU usage %, # Memory, Memory Tiering Type, Memory usage %, Console, # NICs, # HBAs, # VMs total, # VMs, VMs per Core, # vCPUs, vCPUs per Core, vRAM, VM Used memory, VM Memory Swapped, VM Memory Ballooned, VMotion support, Storage VMotion support, Current EVC, Max EVC, Assigned License(s), ATS Heartbeat, ATS Locking, Current CPU power man. policy, Supported CPU power man., Host Power Policy, ESX Version, Boot time, DNS Servers, DHCP, Domain, Domain List, DNS Search Order, NTP Server(s), NTPD running, Time Zone, Time Zone Name, GMT Offset, Vendor, Model, Serial number, Service tag, OEM specific string, BIOS Vendor, BIOS Version, BIOS Date, Certificate Issuer, Certificate Start Date, Certificate Expiry Date, Certificate Status, Certificate Subject, AutoDeploy.MachineIdentity, UUID, Host, LOCATION
```


### vHBA

**RVTools columns:**

```
Host, Datacenter, Cluster, Device, Type, Status, Bus, Pci, Driver, Model, WWN
```


### vNIC

**RVTools columns:**

```
Host, Datacenter, Cluster, Network Device, Driver, Speed, Duplex, MAC, Switch, Uplink port, PCI, WakeOn
```


### vSwitch

**RVTools columns:**

```
Host, Datacenter, Cluster, Switch, # Ports, Free Ports, Promiscuous Mode, Mac Changes, Forged Transmits, Traffic Shaping, Width, Peak, Burst, Policy, Reverse Policy, Notify Switch, Rolling Order, Offload, TSO, Zero Copy Xmit, MTU
```


### vPort

**RVTools columns:**

```
Host, Datacenter, Cluster, Port Group, Switch, VLAN, Promiscuous Mode, Mac Changes, Forged Transmits, Traffic Shaping, Width, Peak, Burst, Policy, Reverse Policy, Notify Switch, Rolling Order, Offload, TSO, Zero Copy Xmit
```


### dvSwitch

**RVTools columns:**

```
Switch, Datacenter, Name, Vendor, Version, Description, Created, Host members, Max Ports, # Ports, # VMs, In Traffic Shaping, In Avg, In Peak, In Burst, Out Traffic Shaping, Out Avg, Out Peak, Out Burst, CDP Type, CDP Operation, LACP Name, LACP Mode, LACP Load Balance Alg., Max MTU, Contact, Admin Name
```


### dvPort

**RVTools columns:**

```
Port, Switch, Type, # Ports, VLAN, Speed, Full Duplex, Blocked, Allow Promiscuous, Mac Changes, Active Uplink, Standby Uplink, Policy, Forged Transmits, In Traffic Shaping, In Avg, In Peak, In Burst, Out Traffic Shaping, Out Avg, Out Peak, Out Burst, Reverse Policy, Notify Switch, Rolling Order, Check Beacon, Live Port Moving, Check Duplex, Check Error %, Check Speed, Percentage, Block Override, Config Reset, Shaping Override, Vendor Config Override, Sec. Policy Override, Teaming Override, Vlan Override
```


### vSC_VMK

**RVTools columns:**

```
Host, Datacenter, Cluster, Port Group, Device, Mac Address, DHCP, IP Address, IP 6 Address, Subnet mask, Gateway, IP 6 Gateway, MTU
```


### vDatastore

**RVTools columns:**

```
Name, Config status, Address, Accessible, Type, # VMs total, # VMs, Capacity MiB, Provisioned MiB, In Use MiB, Free MiB, Free %, SIOC enabled, SIOC Threshold, # Hosts, Hosts, Cluster name, Cluster capacity MiB, Cluster free space MiB, Block size, Max Blocks, # Extents, Major Version, Version, VMFS Upgradeable, MHA, URL
```


### vMultiPath

> **Not implemented.** Needs HostMultipathInfo / storage path enumeration

**RVTools columns:**

```
Host, Cluster, Datacenter, Datastore, Disk, Display name, Policy, Oper. State, Path 1, Path 1 state, Path 2, Path 2 state, Path 3, Path 3 state, Path 4, Path 4 state, Path 5, Path 5 state, Path 6, Path 6 state, Path 7, Path 7 state, Path 8, Path 8 state, vStorage, Queue depth, Vendor, Model, Revision, Level, Serial #, UUID
```


### vLicense

**RVTools columns:**

```
Name, Key, Labels, Cost Unit, Total, Used, Expiration Date, Features
```


### vFileInfo

> Needs HostDatastoreBrowser + SearchDatastoreSubFolders (datastore file-tree
> walking), a different API area from the property collector. Invar implements
> it in `src-tauri/src/data/vfileinfo.rs`.

**RVTools columns:**

```
Friendly Path Name, File Name, File Type, File Size in bytes, Path
```

**Invar diverges here, deliberately.** It adds a `Modification` column between
`File Size in bytes` and `Path`, because the search already asks for
`modification` and a file listing without a timestamp answers half the question
anyone opens it for. This is the one sheet where a column exists that RVTools
does not have, so a consumer reading this sheet **by column index** will be one
off from `Path` onward. Read by label. RVTools also carries an `Internal Sort
Column` here that is an implementation detail of its own grid; Invar has no
equivalent and does not fake one.


### vHealth

> **Correction (verified 2026-08-31).** The note that this needs
> EventManager / AlarmManager aggregation is **wrong**. Every row in the
> reference export is a check RVTools computes itself from inventory it has
> already collected — there is not a single vCenter alarm in the sheet. The 43
> rows break down as: Foldername 10, CDROM 10, Zombie 6, NTP 5, NTPD 5,
> Snapshot 5, Performance tip 2.

**RVTools columns:**

```
Name, Message, Message type
```

**Checks and their exact wording**, read out of the reference export:

| Type | Message | Source | Implemented |
|---|---|---|---|
| `NTP` | `NTP Server value is null!` | `config.dateTimeInfo.ntpConfig.server` empty | yes |
| `NTPD` | `NTPD service is not running!` | `config.service.service` key `ntpd` | yes |
| `Foldername` | `Inconsistent Foldername! VMname = {vm} Foldername = {folder}` | folder in `config.files.vmPathName` vs VM name, **case-sensitive** | yes |
| `CDROM` | `VM has a CDROM device connected! {label}` | `VirtualCdrom` with `connectable.connected` | yes |
| `Snapshot` | `VM has an active snapshot! {name} created on {yyyy/MM/dd HH:mm:ss}` | `snapshot.rootSnapshotList`, flattened | yes |
| `Zombie` | `Possibly a Zombie vmdk file! Please check.` | datastore file walk vs registered disks | **no** — needs `HostDatastoreBrowser` / `SearchDatastoreSubFolders` |
| `Performance tip` | `In-Memory VM performance improvement possible! Please check documentation` | unknown | **no** — trigger not determinable from the export |

RVTools emits host findings before VM findings, and groups a VM's findings
together; we do the same.


### vMetaData

**RVTools columns:**

```
RVTools major version, RVTools version, xlsx creation datetime, Server
```

