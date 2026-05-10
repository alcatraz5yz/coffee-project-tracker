const projects = [
  {
    id: "EF1157",
    name: "EF1157 CoffeeB Pluto",
    family: "CoffeeB M4",
    market: "EU / CH",
    variantGroup: "Pluto platform",
    owner: "PCS / Approbation",
    phase: "IEC approval package",
    target: "TS2 / VDE package",
    updated: "May 9, 2026",
    health: "Watch",
    progress: 62,
    closeout: {
      status: "Bei VDE",
      summary: "VDE package is in work; final CB/Safety/EMC/ErP reports and closing document set are not confirmed yet.",
      gates: [
        { label: "VDE Unterlagen eingereicht", status: "Done" },
        { label: "Pruefungen bestanden", status: "Open" },
        { label: "Finaler VDE/CB Bericht vorhanden", status: "Blocked" },
        { label: "EMC / ErP Reports vorhanden", status: "Blocked" },
        { label: "Keine offenen VDE Rueckfragen", status: "Open" },
        { label: "Finale Dokumente abgelegt", status: "Open" }
      ]
    },
    fachfreigabe: {
      gates: [
        { label: "VDE bestanden" },
        { label: "Finaler VDE/CB Bericht akzeptiert" },
        { label: "EMC / ErP Berichte akzeptiert" },
        { label: "Keine offenen VDE Rueckfragen" },
        { label: "Ziffern fachlich geprueft" },
        { label: "Finale Dokumente vollstaendig" },
        { label: "PCS Abschluss bestaetigt" }
      ]
    },
    stats: [
      ["Certification", "68%", "Safety, EMC, ErP package in progress"],
      ["Approbation", "11/18", "Click to review Ziffer checklist"],
      ["Builds", "TS1 done", "OOT, PT1, PT2, TS1 tracked; TS2 pending"],
      ["Electronics", "Rev B", "Mainboard and MMI schematics available"],
      ["PCS", "Open", "Need final evidence table and samples"]
    ],
    subtopics: {
      Approbation: {
        summary: "IEC / VDE approval checklist by Ziffer. Status values are a first draft from the EF1157 folder and should be corrected as the team reviews evidence.",
        ziffern: [
          { nr: "4", title: "General conditions / product scope", status: "Done", owner: "Approbation", evidence: "Order forms, product type labels, user manual draft", evidenceLinks: [
            { label: "Approbation order form", href: "evidence-1157/Zulassungen/IEC/05 Pflichtenheft Offerte Auftrag/Approbationsauftrag/EF1157_Orderform Approbation_v.1.0-en_V01_20251016.xlsx" },
            { label: "VDE order PDF", href: "evidence-1157/Zulassungen/IEC/05 Pflichtenheft Offerte Auftrag/Offerte Aufrag/Order-338940-13248912.PDF" },
            { label: "User manual draft", href: "evidence-1157/Zulassungen/IEC/03 User Manual/EF1157 UM  CoffeeB Pluto_Draft.pdf" }
          ], note: "Project scope exists; final scope confirmation still useful." },
          { nr: "7", title: "Marking and instructions", status: "Open", owner: "Approbation", evidence: "03 User Manual, 04 Typenschild", evidenceLinks: [
            { label: "User manual draft", href: "evidence-1157/Zulassungen/IEC/03 User Manual/EF1157 UM  CoffeeB Pluto_Draft.pdf" },
            { label: "IEC/DE/CH type labels", href: "evidence-1157/Zulassungen/IEC/04 Typenschild/EF1157 Zulassungstypenschilder Delica 220-240V   IEC_DE_CH_DE  ink EF 1157.docx" },
            { label: "Brazil 220V 60Hz label", href: "evidence-1157/Zulassungen/IEC/04 Typenschild/EF1157 Zulassungstypenschilde 220V 60 Hz  BR.docx" }
          ], note: "Need final rating plate and manual consistency check." },
          { nr: "8", title: "Protection against live parts", status: "Open", owner: "Safety", evidence: "Photos / construction review", evidenceLinks: [
            { label: "TS1 photo folder", href: "evidence-1157/Zulassungen/IEC/13 Fotos/Testserie 1/" },
            { label: "Wiring diagram", href: "evidence-1157/Zulassungen/IEC/08 Schema Gerät/Elektroschema/1157A801-w00.02.00.Wiring Diagramm.pdf" }
          ], note: "Needs explicit evidence set." },
          { nr: "10", title: "Input and current", status: "Done", owner: "PCS", evidence: "12 Untersuchungen / Messresultate 220V/230V/240V", evidenceLinks: [
            { label: "230V Ziffer 11 measurement", href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Messresultate/EF1157_ZIFFER11_230V_50HZ_V2.XLS" },
            { label: "220V 60Hz measurement", href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Messresultate/EF1157_ZIFFER11_220V_60HZ.XLS" },
            { label: "Ziffer 10 reference measurement", href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Messresultate/EF1110 messung § 10 Mittelwert.xlsx" }
          ], note: "Measurement files are present." },
          { nr: "11", title: "Heating", status: "Done", owner: "PCS", evidence: "Prototype measurements, heater / TCO photos", evidenceLinks: [
            { label: "OOT Ziffer 11 measurement", href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Prototype/Messresultate/EF1157_OOT_041 ZIFFER11.XLS" },
            { label: "230V Ziffer 11 measurement", href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Messresultate/EF1157_ZIFFER11_230V_50HZ_V2.XLS" },
            { label: "Heater/TCO photo folder", href: "evidence-1157/Zulassungen/IEC/13 Fotos/Testserie 1/Punktschweossung Heizer _TCO/" }
          ], note: "Folder contains heater/TCO measurement evidence." },
          { nr: "13", title: "Leakage current / electric strength", status: "Open", owner: "Safety", evidence: "Pre-approval report draft", evidenceLinks: [
            { label: "Draft report folder", href: "evidence-1157/Zulassungen/IEC/15 CB Prüfberichte Safety EMC ErP/Draft Report/" },
            { label: "TS1 type-test draft", href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/TS1-Serie/xxxx_EF1157  VDE 338940 UB Typenprüfung.docm" }
          ], note: "Need final result mapping." },
          { nr: "15", title: "Moisture resistance", status: "Open", owner: "Safety", evidence: "Pre-approval report draft", evidenceLinks: [
            { label: "Draft report folder", href: "evidence-1157/Zulassungen/IEC/15 CB Prüfberichte Safety EMC ErP/Draft Report/" },
            { label: "Fluid system", href: "evidence-1157/Zulassungen/IEC/08 Schema Gerät/Fluidschema/EF1157_EF1156-BU_Fluidsystem.pdf" }
          ], note: "Check if test is done or not applicable." },
          { nr: "16", title: "Leakage after moisture", status: "Open", owner: "Safety", evidence: "Pre-approval report draft", evidenceLinks: [
            { label: "Draft report folder", href: "evidence-1157/Zulassungen/IEC/15 CB Prüfberichte Safety EMC ErP/Draft Report/" },
            { label: "Fluid system", href: "evidence-1157/Zulassungen/IEC/08 Schema Gerät/Fluidschema/EF1157_EF1156-BU_Fluidsystem.pdf" }
          ], note: "Needs final report status." },
          { nr: "17", title: "Overload protection", status: "Not needed", owner: "Approbation", evidence: "To confirm", note: "Placeholder until standard matrix is reviewed." },
          { nr: "19", title: "Abnormal operation / motor heating", status: "Done", owner: "PCS", evidence: "Ziffer 19 measurements, photos, fault imitation sheet", evidenceLinks: [
            { label: "Fault imitation sheet", href: "evidence-1157/Zulassungen/IEC/07 Elektronik/EF1157-V003HW_V00-47SW_230V_Fault Imitation Sheet_EN.pdf" },
            { label: "Pump Ziffer 19 XLS", href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Prototype/Messresultate/EF1157_ZIFFER19_PUMPE_19_11_2.XLS" },
            { label: "Ziffer 19 photo folder", href: "evidence-1157/Zulassungen/IEC/13 Fotos/Testserie 1/Messungen/Ziffer_19_4/" }
          ], note: "Folder explicitly contains Ziffer 19 measurement evidence." },
          { nr: "20", title: "Stability and mechanical hazards", status: "Open", owner: "Mechanical", evidence: "Prototype photos", evidenceLinks: [
            { label: "TS1 photo folder", href: "evidence-1157/Zulassungen/IEC/13 Fotos/Testserie 1/" },
            { label: "Views/colors folder", href: "evidence-1157/Zulassungen/IEC/13 Fotos/Testserie 1/Ansichten_Farben/" }
          ], note: "Need mechanical checklist." },
          { nr: "22", title: "Construction", status: "Open", owner: "Safety", evidence: "Schematics, wiring diagram, fluid system", evidenceLinks: [
            { label: "Mainboard schematic", href: "evidence-1157/Zulassungen/IEC/07 Elektronik/Elektronik  BOM_Schema_Layout/Elektonink BOM_Schema_Layout/EF1157-Mainboard-SCH-003-B-V1-CE.PDF" },
            { label: "MMI schematic", href: "evidence-1157/Zulassungen/IEC/07 Elektronik/Elektronik  BOM_Schema_Layout/Elektonink BOM_Schema_Layout/EF1157-MMI-SCH-002-B-V1.PDF" },
            { label: "Wiring diagram", href: "evidence-1157/Zulassungen/IEC/08 Schema Gerät/Elektroschema/1157A801-w00.02.00.Wiring Diagramm.pdf" },
            { label: "Fluid system", href: "evidence-1157/Zulassungen/IEC/08 Schema Gerät/Fluidschema/EF1157_EF1156-BU_Fluidsystem.pdf" }
          ], note: "Needs final construction review." },
          { nr: "24", title: "Components", status: "Done", owner: "Approbation", evidence: "10 Komponenten, 11 Materialpruefungen, approvals", evidenceLinks: [
            { label: "Component evidence folder", href: "evidence-1157/Zulassungen/IEC/10 Komponenten/Komponenten Nachweise/" },
            { label: "Pump VDE marks approval", href: "evidence-1157/Zulassungen/IEC/10 Komponenten/0155224 _0152436  Pump Sysko SPX.H122A/VDE_Marks_Approval_40060855_300.pdf" },
            { label: "TCO pump approval", href: "evidence-1157/Zulassungen/IEC/10 Komponenten/0153271 TCO Pump  SF94R0/VDE 40035880.pdf" }
          ], note: "Many component certificates are present." },
          { nr: "25", title: "Supply connection / external cords", status: "Done", owner: "Approbation", evidence: "Netzkabel EU/CH component approvals", evidenceLinks: [
            { label: "EU cord VDE", href: "evidence-1157/Zulassungen/IEC/10 Komponenten/Komponenten Nachweise/01556777  Netzkabel EU SF-71/0054494  H05VV-F 3G0.75mm2/VDE 120557.pdf" },
            { label: "CH cord VDE", href: "evidence-1157/Zulassungen/IEC/10 Komponenten/Komponenten Nachweise/0156693 Netzkabel CH SF 285/0054494 Cord H05VV-F3G 0.75/VDE 120557.pdf" },
            { label: "Alternative EU cord", href: "evidence-1157/Zulassungen/IEC/10 Komponenten/Komponenten Nachweise/0156694 Netzkabel EU  CW3191/0106886_cord H05VV-F3G 0.75/VDE 109835.pdf" }
          ], note: "Cord approval folders are present." },
          { nr: "27", title: "Earthing provision", status: "Blocked", owner: "Safety", evidence: "Wiring diagram", evidenceLinks: [
            { label: "Wiring diagram", href: "evidence-1157/Zulassungen/IEC/08 Schema Gerät/Elektroschema/1157A801-w00.02.00.Wiring Diagramm.pdf" },
            { label: "Mainboard schematic", href: "evidence-1157/Zulassungen/IEC/07 Elektronik/Elektronik  BOM_Schema_Layout/Elektonink BOM_Schema_Layout/EF1157-Mainboard-SCH-003-B-V1-CE.PDF" }
          ], note: "Need confirm appliance class and PE concept." },
          { nr: "29", title: "Clearances / creepage distances", status: "Open", owner: "HW", evidence: "Mainboard PCB/artwork PDFs", evidenceLinks: [
            { label: "Mainboard PCB", href: "evidence-1157/Zulassungen/IEC/07 Elektronik/Elektronik  BOM_Schema_Layout/Elektonink BOM_Schema_Layout/EF1157-Mainboard-PCB-003-A.PDF" },
            { label: "Mainboard final artworks", href: "evidence-1157/Zulassungen/IEC/07 Elektronik/Elektronik  BOM_Schema_Layout/Elektonink BOM_Schema_Layout/EF1157-Mainboard-Final Artworks-003-A.PDF" },
            { label: "PCB supplier certificates", href: "evidence-1157/Zulassungen/IEC/07 Elektronik/Leiterplatten Hersteller/" }
          ], note: "Needs PCB review record." },
          { nr: "30", title: "Resistance to heat and fire", status: "Done", owner: "Materials", evidence: "13 Fotos/Testserie 1/Messungen/Ziffer_30", evidenceLinks: [
            { label: "Ziffer 30 photo folder", href: "evidence-1157/Zulassungen/IEC/13 Fotos/Testserie 1/Messungen/Ziffer_30/" },
            { label: "Draft Table 30", href: "evidence-1157/Zulassungen/IEC/11 Materialprüfungen/Tabelle 30/Draft_EF1157 Tab. 30_VDE.docx" },
            { label: "Draft material tests", href: "evidence-1157/Zulassungen/IEC/11 Materialprüfungen/GWT interne  Prüfungen/Draft_EF1157_VDE 338940_Materialprüfungen GWT  24.02.2026.docx" }
          ], note: "Ziffer_30 folder exists." },
          { nr: "32", title: "Radiation / toxicity / similar hazards", status: "Not needed", owner: "Approbation", evidence: "To confirm", note: "Likely not applicable, confirm in matrix." }
        ]
      }
    },
    certification: [
      { name: "IEC 60335-1 / 2-15 review", state: "In review", done: false },
      { name: "VDE order / approval package", state: "Started", done: true },
      { name: "Safety fault imitation sheet", state: "Available", done: true },
      { name: "EMC evidence", state: "Pending", done: false },
      { name: "ErP evidence", state: "Pending", done: false },
      { name: "Declaration / conformity docs", state: "Draft", done: false }
    ],
    builds: [
      { label: "OOT", date: "2025-10", state: "Done", samples: "3 pcs", note: "Out-of-tool parts checked for fit, leakage, and assembly basics." },
      { label: "PT1", date: "2025-11", state: "Done", samples: "6 pcs", note: "First prototype build used for basic function and early measurements." },
      { label: "PT2", date: "2025-12", state: "Done", samples: "8 pcs", note: "Pre-approval and measurement photos present." },
      { label: "TS1", date: "2026-01", state: "Done", samples: "12 pcs", note: "Dispatch lists and test photos present." },
      { label: "TS2", date: "2026-03", state: "Open", samples: "Planned", note: "Dispatch list exists, evidence still needs review." },
      { label: "PVT", date: "2026-05", state: "Not started", samples: "TBD", note: "Production validation scope not yet frozen." }
    ],
    risks: [
      { level: "High", text: "Mains heater/pump safety path must stay hardware-protected." },
      { level: "Medium", text: "MMI replacement is not a drop-in display; original MMI is 2 buttons plus LEDs." },
      { level: "Medium", text: "Need final mapping between build samples and approval documents." }
    ],
    tasks: [
      { area: "Certification", task: "Confirm latest VDE order scope", owner: "Approbation", due: "2026-05-14", status: "Open", builds: "Alle" },
      { area: "Electronics", task: "Review EF1157 mainboard Rev B against wiring diagram", owner: "HW", due: "2026-05-16", status: "Open", builds: "PT1" },
      { area: "Build", task: "Link TS1 / TS2 dispatch lists to actual samples", owner: "PCS", due: "2026-05-17", status: "Open", builds: "TS1,TS2" },
      { area: "PCS", task: "Create sample evidence matrix", owner: "PCS", due: "2026-05-20", status: "In progress", builds: "PT1,OOT,TS1,TS2" },
      { area: "Approbation", task: "Check safety, EMC, ErP report completeness", owner: "Approbation", due: "2026-05-22", status: "Open", builds: "TS1" },
      { area: "Certification", task: "Prepare final conformity document list", owner: "Approbation", due: "2026-05-28", status: "Blocked", builds: "TS2" },
      { area: "Approbation", task: "Clarify missing final CB / EMC / ErP report package", owner: "Approbation", due: "2026-05-29", status: "Blocked", builds: "TS2" },
      { area: "Build", task: "Create clean TS1 / TS2 sample and dispatch overview", owner: "PCS", due: "2026-05-30", status: "Open", builds: "TS1,TS2" },
      { area: "Certification", task: "Confirm country implementation folder scope", owner: "Approbation", due: "2026-06-03", status: "Open", builds: "TS1,TS2" }
    ],
    documents: [
      { name: "User manual draft", path: "03 User Manual/EF1157 UM CoffeeB Pluto_Draft.pdf", state: "Available" },
      { name: "Wiring diagram", path: "08 Schema Geraet/Elektroschema/1157A801-w00.02.00.Wiring Diagramm.pdf", state: "Available" },
      { name: "Fluid system", path: "08 Schema Geraet/Fluidschema/EF1157_EF1156-BU_Fluidsystem.pdf", state: "Available" },
      { name: "Mainboard schematic", path: "07 Elektronik/.../EF1157-Mainboard-SCH-003-B-V1-CE.PDF", state: "Available" },
      { name: "MMI schematic", path: "07 Elektronik/.../EF1157-MMI-SCH-002-B-V1.PDF", state: "Available" },
      { name: "Fault imitation sheet", path: "07 Elektronik/EF1157-V003HW_V00-47SW_230V_Fault Imitation Sheet_EN.pdf", state: "Available" },
      { name: "Approval BOM", path: "06 Materialliste Explo/EF1157_Approbation BOM_V0.0_20251212.xlsm", state: "Available" },
      { name: "Final CB / VDE report", path: "15 CB Pruefberichte Safety EMC ErP", state: "Needs check" }
    ],
    documentGroups: [
      {
        area: "Administration",
        status: "Available",
        count: "13 files",
        summary: "Project plan, dispatch lists, delivery note templates, minutes, and project email.",
        primary: "Project master plan and TS1/TS2 dispatch lists",
        href: "evidence-1157/Zulassungen/IEC/01 Administration/"
      },
      {
        area: "Standards / Changes",
        status: "Available",
        count: "6 files",
        summary: "IEC/EN 60335 reference standards and change documents.",
        primary: "IEC 60335-1, IEC 60335-2-15, EN 60335-1",
        href: "evidence-1157/Zulassungen/IEC/02 Änderungen/"
      },
      {
        area: "Manual / Labels",
        status: "Open",
        count: "5 files",
        summary: "User manual draft plus EU/CH/DE and Brazil rating label versions.",
        primary: "Manual draft and type labels",
        href: "evidence-1157/Zulassungen/IEC/03 User Manual/EF1157 UM  CoffeeB Pluto_Draft.pdf"
      },
      {
        area: "Order / Scope",
        status: "Available",
        count: "20 files",
        summary: "VDE offers, order, approbation order form, forms, product specification references, and electronics package duplicate.",
        primary: "Order 338940 and EF1157 approbation order form",
        href: "evidence-1157/Zulassungen/IEC/05 Pflichtenheft Offerte Auftrag/"
      },
      {
        area: "BOM / Exploded Views",
        status: "Available",
        count: "4 files",
        summary: "Approval BOM, exploded assembly drawings, and brewing unit exploded assembly.",
        primary: "EF1157 Approbation BOM and exploded assembly",
        href: "evidence-1157/Zulassungen/IEC/06 Materialliste Explo/"
      },
      {
        area: "Electronics",
        status: "Available",
        count: "116 files",
        summary: "Mainboard/MMI schematics, BOMs, PCB artwork, fault imitation sheet, triac/capacitor/varistor/connector/PCB supplier approvals.",
        primary: "Mainboard + MMI electronics package",
        href: "evidence-1157/Zulassungen/IEC/07 Elektronik/"
      },
      {
        area: "Device Schemas",
        status: "Available",
        count: "2 files",
        summary: "Wiring diagram and fluid system diagram.",
        primary: "Electrical and fluid schemas",
        href: "evidence-1157/Zulassungen/IEC/08 Schema Gerät/"
      },
      {
        area: "Bautelliste",
        status: "Open",
        count: "4 real versions",
        summary: "Internal and VDE part-list versions from 16 Feb, 18 Feb, 7 Mar, and 20 Apr 2026.",
        primary: "Latest internal Bautelliste 20.04.2026",
        href: "evidence-1157/Zulassungen/IEC/09 Bautelliste/"
      },
      {
        area: "Components",
        status: "Available",
        count: "104 files",
        summary: "Component certificates including cords, pumps, heaters, flowmeter, STB/TCO, capacitors, varistors, and related VDE approvals.",
        primary: "Komponenten Nachweise",
        href: "evidence-1157/Zulassungen/IEC/10 Komponenten/"
      },
      {
        area: "Material / GWT",
        status: "Open",
        count: "54 files",
        summary: "GWT evidence, Table 30 draft, UL yellow cards, material approvals, and EF1157 material-test draft.",
        primary: "Table 30 and material test drafts",
        href: "evidence-1157/Zulassungen/IEC/11 Materialprüfungen/"
      },
      {
        area: "Investigations",
        status: "Current",
        count: "61 files",
        summary: "Approbation Word reports, measurements, Ziffer 11/19 data, prototype evidence, and TS1 type-test draft.",
        primary: "12 Untersuchungen reports and measurement evidence",
        href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/"
      },
      {
        area: "Photos",
        status: "Available",
        count: "104 files",
        summary: "Prototype and TS1 photos, measurement photos, Ziffer 19 and Ziffer 30 photo sets, heater/TCO weld photos.",
        primary: "Testserie 1 photo evidence",
        href: "evidence-1157/Zulassungen/IEC/13 Fotos/Testserie 1/"
      },
      {
        area: "PAK / LFGB",
        status: "Open",
        count: "3 files",
        summary: "PAK BOM, PAK list for VDE, and exploded assembly PAK evidence. LFGB/individual test subfolders exist but no files were found in this scan.",
        primary: "EF1157 PAK list and PAK BOM",
        href: "evidence-1157/Zulassungen/IEC/14 PAK Bewertung/"
      },
      {
        area: "CB / Safety / EMC / ErP",
        status: "Blocked",
        count: "3 files",
        summary: "Folder exists but final report package is not populated; only pressure-equipment classification, EF1107 info doc, and one temporary compliance lock file were found.",
        primary: "Final report package still needs confirmation",
        href: "evidence-1157/Zulassungen/IEC/15 CB Prüfberichte Safety EMC ErP/"
      },
      {
        area: "Country Implementation",
        status: "Open",
        count: "0 files",
        summary: "Folder exists but is empty. This matters for variants like EF1234 Brazil.",
        primary: "No country files found yet",
        href: "evidence-1157/Zulassungen/IEC/16 Länderumsetzungen/"
      }
    ],
    reportVersions: [
      {
        project: "EF1157",
        build: "PT / Pre-Approval",
        version: "Original archive",
        modified: "2026-01-27 10:37",
        size: "18.0 MB",
        state: "Archived",
        file: "EF1157 UB Pre-Approval PT.docx",
        href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Prototype/Archiv/EF1157 UB Pre-Approval PT.docx"
      },
      {
        project: "EF1157",
        build: "PT / Pre-Approval",
        version: "REV 02-02-2026",
        modified: "2026-02-02 16:20",
        size: "3.8 MB",
        state: "Archived",
        file: "EF1157 UB Pre-Approval REV 02-02-2026.docx",
        href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Prototype/EF1157 UB Pre-Approval   REV 02-02-2026.docx"
      },
      {
        project: "EF1157",
        build: "PT / Pre-Approval",
        version: "PCS2-E copy",
        modified: "2026-02-02 17:26",
        size: "3.8 MB",
        state: "Archived",
        file: "EF1157 UB Pre-Approval PT REV 02-02-2026 PCS2-E - Kopie.docx",
        href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Prototype/Archiv/EF1157 UB Pre-Approval PT  REV 02-02-2026 PCS2-E - Kopie.docx"
      },
      {
        project: "EF1157",
        build: "PT / Pre-Approval",
        version: "REV 02-02-2026 PCS2-E",
        modified: "2026-02-04 11:53",
        size: "3.8 MB",
        state: "Current",
        file: "EF1157 UB Pre-Approval PT REV 02-02-2026 PCS2-E.docx",
        href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Prototype/EF1157 UB Pre-Approval PT  REV 02-02-2026 PCS2-E.docx"
      },
      {
        project: "EF1157",
        build: "TS1 / Typenpruefung",
        version: "VDE 338940 UB",
        modified: "2026-02-16 14:19",
        size: "14.5 MB",
        state: "Current",
        file: "xxxx_EF1157 VDE 338940 UB Typenpruefung.docm",
        href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/TS1-Serie/xxxx_EF1157  VDE 338940 UB Typenprüfung.docm"
      }
    ]
  },
  {
    id: "EF1234",
    name: "EF1234 CoffeeB Pluto Brazil",
    family: "CoffeeB M4",
    market: "Brazil",
    variantGroup: "Pluto platform",
    variantOf: "EF1157",
    owner: "PCS / Approbation",
    phase: "Brazil variant definition",
    target: "INMETRO / Brazil launch package",
    updated: "May 9, 2026",
    health: "Watch",
    progress: 38,
    closeout: {
      status: "In Arbeit",
      summary: "Brazil variant is not ready for VDE/INMETRO closeout; delta scope and local documents are still open.",
      gates: [
        { label: "Delta Scope EF1157 -> EF1234 bestaetigt", status: "Open" },
        { label: "Brazil Unterlagen eingereicht", status: "Open" },
        { label: "Pruefungen bestanden", status: "Open" },
        { label: "Finale Brazil Dokumente abgelegt", status: "Blocked" }
      ]
    },
    fachfreigabe: {
      gates: [
        { label: "Delta-Scope EF1157 zu EF1234 bestaetigt" },
        { label: "Brazil-Unterlagen fachlich geprueft" },
        { label: "Pruefungen bestanden" },
        { label: "Finale Brazil Dokumente vollstaendig" },
        { label: "PCS Abschluss bestaetigt" }
      ]
    },
    stats: [
      ["Certification", "35%", "Brazil delta package needs owner confirmation"],
      ["Approbation", "4/18", "Brazil checklist based on EF1157 carry-over"],
      ["Builds", "PT1 planned", "Needs variant samples and Brazil-specific labels"],
      ["Electronics", "Carry-over", "Check mains, plug, label, and approvals delta"],
      ["PCS", "Open", "Need Brazil evidence matrix"]
    ],
    subtopics: {
      Approbation: {
        summary: "Brazil variant checklist. Most rows should carry over from EF1157 only after the delta review confirms labels, mains data, plug/cord, documentation language, and local approval requirements.",
        ziffern: [
          { nr: "4", title: "General conditions / product scope", status: "Open", owner: "Approbation", evidence: "Variant definition", note: "Confirm EF1234 scope versus EF1157." },
          { nr: "7", title: "Marking and instructions", status: "Open", owner: "Approbation", evidence: "Brazil label and manual", note: "Portuguese manual and Brazil rating plate required." },
          { nr: "8", title: "Protection against live parts", status: "Not needed", owner: "Safety", evidence: "EF1157 carry-over", note: "Carry-over if housing and construction stay identical." },
          { nr: "10", title: "Input and current", status: "Open", owner: "PCS", evidence: "Brazil mains measurement", note: "Confirm local mains variant and rating." },
          { nr: "11", title: "Heating", status: "Open", owner: "PCS", evidence: "Variant thermal measurements", note: "Repeat if heater, mains, or control parameters differ." },
          { nr: "13", title: "Leakage current / electric strength", status: "Open", owner: "Safety", evidence: "Brazil approval package", note: "Needs local approval evidence." },
          { nr: "15", title: "Moisture resistance", status: "Not needed", owner: "Safety", evidence: "EF1157 carry-over", note: "Carry-over only if construction unchanged." },
          { nr: "16", title: "Leakage after moisture", status: "Not needed", owner: "Safety", evidence: "EF1157 carry-over", note: "Carry-over only if test house accepts it." },
          { nr: "17", title: "Overload protection", status: "Not needed", owner: "Approbation", evidence: "EF1157 matrix", note: "Confirm with local standard matrix." },
          { nr: "19", title: "Abnormal operation / motor heating", status: "Open", owner: "PCS", evidence: "EF1157 test plus delta note", note: "Confirm pump/heater/control carry-over." },
          { nr: "20", title: "Stability and mechanical hazards", status: "Not needed", owner: "Mechanical", evidence: "EF1157 carry-over", note: "Carry-over if mechanical design unchanged." },
          { nr: "22", title: "Construction", status: "Open", owner: "Safety", evidence: "Delta construction review", note: "Document all Brazil-specific differences." },
          { nr: "24", title: "Components", status: "Open", owner: "Approbation", evidence: "Brazil BOM delta", note: "Check plug, cord, label, PSU/heater variants." },
          { nr: "25", title: "Supply connection / external cords", status: "Blocked", owner: "Approbation", evidence: "Brazil cord/plug approval", note: "Needs Brazil-specific cord and plug decision." },
          { nr: "27", title: "Earthing provision", status: "Open", owner: "Safety", evidence: "Wiring diagram delta", note: "Confirm appliance class and PE concept." },
          { nr: "29", title: "Clearances / creepage distances", status: "Not needed", owner: "HW", evidence: "EF1157 PCB carry-over", note: "Carry-over if PCB and voltage class unchanged." },
          { nr: "30", title: "Resistance to heat and fire", status: "Open", owner: "Materials", evidence: "Material carry-over list", note: "Confirm materials and suppliers unchanged." },
          { nr: "32", title: "Radiation / toxicity / similar hazards", status: "Not needed", owner: "Approbation", evidence: "EF1157 matrix", note: "Confirm in local matrix." }
        ]
      }
    },
    certification: [
      { name: "Brazil approval route / INMETRO scope", state: "Open", done: false },
      { name: "Portuguese user manual", state: "Open", done: false },
      { name: "Brazil rating plate and packaging labels", state: "Open", done: false },
      { name: "EF1157 carry-over justification", state: "Started", done: true },
      { name: "Cord / plug component approvals", state: "Blocked", done: false }
    ],
    builds: [
      { label: "OOT", date: "2026-02", state: "Not needed", samples: "Carry-over", note: "Use EF1157 OOT if plastic/tooling is identical." },
      { label: "PT1", date: "2026-05", state: "Open", samples: "4 pcs", note: "Build Brazil variant samples with local label, cord, and firmware data." },
      { label: "PT2", date: "2026-06", state: "Planned", samples: "6 pcs", note: "Approval samples after PT1 delta feedback." },
      { label: "TS1", date: "2026-07", state: "Planned", samples: "10 pcs", note: "Tooling/series-intent Brazil confirmation." },
      { label: "TS2", date: "2026-08", state: "Planned", samples: "TBD", note: "Final evidence build for launch release." }
    ],
    risks: [
      { level: "High", text: "Brazil plug, cord, label, language, and local approval route are not yet closed." },
      { level: "Medium", text: "Carry-over from EF1157 must be documented row by row, not assumed globally." }
    ],
    tasks: [
      { area: "Certification", task: "Confirm Brazil approval route and required standards", owner: "Approbation", due: "2026-05-15", status: "Open", builds: "Alle" },
      { area: "Build", task: "Define PT1 Brazil sample configuration", owner: "PCS", due: "2026-05-18", status: "Open", builds: "PT1" },
      { area: "Approbation", task: "Create EF1157 to EF1234 delta matrix", owner: "Approbation", due: "2026-05-20", status: "Open", builds: "PT1,TS1" },
      { area: "Electronics", task: "Check mains, cord, plug, and PCB carry-over", owner: "HW", due: "2026-05-22", status: "Blocked", builds: "PT1,TS1" },
      { area: "PCS", task: "Link Brazil samples to evidence package", owner: "PCS", due: "2026-05-28", status: "Open", builds: "PT1,TS1" }
    ],
    documents: [
      { name: "EF1157 carry-over baseline", path: "EF1157 approval package", state: "Available" },
      { name: "Brazil delta matrix", path: "To create", state: "Needs check" },
      { name: "Portuguese manual", path: "To create", state: "Open" },
      { name: "Brazil rating plate", path: "To create", state: "Open" },
      { name: "Cord / plug approval", path: "To define", state: "Blocked" }
    ],
    reportVersions: [
      {
        project: "EF1234",
        build: "127V PT1 Brazil",
        version: "PCS8 / 30-04-2026",
        modified: "2026-05-03 17:08",
        size: "8.0 MB",
        state: "Current",
        file: "draft_ef1234_127v_60hz_22042026_rev_pcs8_30042026.docm",
        href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/127V pt1 EF-1234/draft_ef1234_127v_60hz_22042026_rev_pcs8_30042026.docm"
      }
    ]
  },
  {
    id: "EF1107",
    name: "EF1107 CoffeeB Bale",
    family: "CoffeeB M4",
    owner: "Approbation",
    phase: "Reference project",
    target: "Carry-over checks",
    updated: "May 9, 2026",
    health: "Good",
    progress: 78,
    closeout: {
      status: "Finale Dokumente fehlen",
      summary: "Reference project is mostly usable, but carry-over applicability still needs explicit confirmation.",
      gates: [
        { label: "Reference VDE Dokumente vorhanden", status: "Done" },
        { label: "Carry-over bestaetigt", status: "Open" },
        { label: "Finale Vergleichsnotiz abgelegt", status: "Open" }
      ]
    },
    fachfreigabe: {
      gates: [
        { label: "Reference-Dokumente vollstaendig" },
        { label: "Carry-over fachlich bestaetigt" },
        { label: "PCS Abschluss bestaetigt" }
      ]
    },
    stats: [
      ["Certification", "82%", "Reference documents available"],
      ["Approbation", "3/6", "Reference checklist placeholder"],
      ["Builds", "Stable", "Used as comparison base"],
      ["Electronics", "Carry-over", "Check shared components"],
      ["PCS", "Open", "Sample mapping needed"]
    ],
    subtopics: {
      Approbation: {
        summary: "Reference project checklist for comparing carry-over approvals.",
        ziffern: [
          { nr: "7", title: "Marking and instructions", status: "Done", owner: "Approbation", evidence: "Reference forms", note: "Available as carry-over input." },
          { nr: "19", title: "Abnormal operation", status: "Open", owner: "PCS", evidence: "Reference test report", note: "Needs comparison." },
          { nr: "24", title: "Components", status: "Open", owner: "Approbation", evidence: "Shared component list", note: "Check common parts." },
          { nr: "29", title: "Clearances / creepage", status: "Not needed", owner: "HW", evidence: "Reference only", note: "Only relevant if PCB reused." },
          { nr: "30", title: "Heat and fire", status: "Done", owner: "Materials", evidence: "Reference material evidence", note: "Available." },
          { nr: "32", title: "Radiation / toxicity", status: "Not needed", owner: "Approbation", evidence: "Reference matrix", note: "Confirm." }
        ]
      }
    },
    certification: [
      { name: "Reference VDE documents", state: "Available", done: true },
      { name: "Carry-over applicability", state: "In review", done: false },
      { name: "Shared component approvals", state: "Partial", done: false }
    ],
    builds: [
      { label: "Reference", date: "2025", state: "Available", note: "Use as baseline for EF1157" }
    ],
    risks: [
      { level: "Medium", text: "Carry-over assumptions must be explicitly confirmed." }
    ],
    tasks: [
      { area: "Approbation", task: "Compare shared approval documents with EF1157", owner: "Approbation", due: "2026-05-21", status: "Open", builds: "PT1,TS1" }
    ],
    documents: [
      { name: "Product specification", path: "05 Pflichtenheft.../230828_1107_product specification...", state: "Available" },
      { name: "Reference forms", path: "05 Pflichtenheft.../Formulare", state: "Available" }
    ],
    reportVersions: [
      {
        project: "EF1107",
        build: "OOT template/reference",
        version: "Template",
        modified: "2026-02-03 12:41",
        size: "81.6 MB",
        state: "Reference",
        file: "Vorlage 7896 EF1107 UB Pre-Approval OOT.docx",
        href: "evidence-1157/Zulassungen/IEC/12 Untersuchungen/Prototype/Vorlage 7896 EF1107 UB Pre-Approval OOT.docx"
      }
    ]
  }
];
