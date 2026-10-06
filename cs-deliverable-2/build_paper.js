// Builds the CS Deliverable 2 manuscript in IEEE conference format (US Letter, two columns).
// Usage: NODE_PATH=$(npm root -g) node build_paper.js
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, SectionType, ShadingType, VerticalAlign,
} = require('docx');

const FONT = 'Times New Roman';
const COL_W = 5040; // 3.5 in column width in DXA

// ---------------------------------------------------------------- references
const REFS = {
  mergel2019: 'I. Mergel, N. Edelmann, and N. Haug, “Defining digital transformation: Results from expert interviews,” *Gov. Inf. Q.*, vol. 36, no. 4, Art. no. 101385, 2019, doi: 10.1016/j.giq.2019.06.002.',
  un2024: 'United Nations DESA, *United Nations E-Government Survey 2024: Accelerating Digital Transformation for Sustainable Development*. New York, NY, USA: UN, 2024.',
  ra12254: 'Congress of the Philippines, *Republic Act No. 12254: E-Governance Act*, 2025, and its Implementing Rules and Regulations, 2026.',
  ra12234: 'Congress of the Philippines, *Republic Act No. 12234: Konektadong Pinoy Act*, 2025, and its Implementing Rules and Regulations, 2025.',
  ra10173: 'Congress of the Philippines, *Republic Act No. 10173: Data Privacy Act of 2012*, 2012.',
  ra11055: 'Congress of the Philippines, *Republic Act No. 11055: Philippine Identification System Act*, 2018.',
  gong2020: 'Y. Gong, J. Yang, and X. Shi, “Towards a comprehensive understanding of digital transformation in government: Analysis of flexibility and enterprise architecture,” *Gov. Inf. Q.*, vol. 37, no. 3, Art. no. 101487, 2020, doi: 10.1016/j.giq.2020.101487.',
  janssen2012: 'M. Janssen, “Sociopolitical aspects of interoperability and enterprise architecture in e-government,” *Soc. Sci. Comput. Rev.*, vol. 30, no. 1, pp. 24–38, 2012, doi: 10.1177/0894439310392187.',
  juraida2024: 'E. Juraida and D. I. Sensuse, “Enterprise architecture as an enabler of digital transformation in the government sector: Success factors and maturity evaluation methodology,” *Eduvest – J. Universal Stud.*, vol. 4, no. 11, pp. 9821–9842, 2024, doi: 10.59188/eduvest.v4i11.43677.',
  togaf: 'The Open Group, *The TOGAF Standard, 10th Edition*. Reading, U.K.: The Open Group, 2022.',
  archimate: 'The Open Group, *ArchiMate 3.2 Specification*. Reading, U.K.: The Open Group, 2023.',
  guijarro2007: 'L. Guijarro, “Interoperability frameworks and enterprise architectures in e-government initiatives in Europe and the United States,” *Gov. Inf. Q.*, vol. 24, no. 1, pp. 89–101, 2007, doi: 10.1016/j.giq.2006.05.003.',
  eif2017: 'European Commission, “New European Interoperability Framework: Promoting seamless services and data flows for European public administrations,” Publications Office of the EU, Luxembourg, 2017.',
  janssen2020: 'M. Janssen, P. Brous, E. Estevez, L. S. Barbosa, and T. Janowski, “Data governance: Organizing data for trustworthy artificial intelligence,” *Gov. Inf. Q.*, vol. 37, no. 3, Art. no. 101493, 2020, doi: 10.1016/j.giq.2020.101493.',
  vandonge2022: 'W. van Donge, N. Bharosa, and M. Janssen, “Data-driven government: Cross-case comparison of data stewardship in data ecosystems,” *Gov. Inf. Q.*, vol. 39, no. 2, Art. no. 101642, 2022, doi: 10.1016/j.giq.2021.101642.',
  xroad: 'Nordic Institute for Interoperability Solutions, “X-Road: Architecture overview,” NIIS, 2024. [Online]. Available: https://x-road.global',
  sousa2019: 'W. G. de Sousa, E. R. P. de Melo, P. H. de S. Bermejo, R. A. S. Farias, and A. O. Gomes, “How and where is artificial intelligence in the public sector going? A literature review and research agenda,” *Gov. Inf. Q.*, vol. 36, no. 4, Art. no. 101392, 2019, doi: 10.1016/j.giq.2019.07.004.',
  straub2023: 'V. J. Straub, D. Morgan, J. Bright, and H. Margetts, “Artificial intelligence in government: Concepts, standards, and a unified framework,” *Gov. Inf. Q.*, vol. 40, no. 4, Art. no. 101881, 2023, doi: 10.1016/j.giq.2023.101881.',
  kankanhalli2019: 'A. Kankanhalli, Y. Charalabidis, and S. Mellouli, “IoT and AI for smart government: A research agenda,” *Gov. Inf. Q.*, vol. 36, no. 2, pp. 304–309, 2019, doi: 10.1016/j.giq.2019.02.003.',
  wirtz2019: 'B. W. Wirtz, J. C. Weyerer, and F. T. Schichtel, “An integrative public IoT framework for smart government,” *Gov. Inf. Q.*, vol. 36, no. 2, pp. 333–345, 2019, doi: 10.1016/j.giq.2018.07.001.',
  abied2022: 'O. Abied, O. Ibrahim, and S. N.-I. M. Kamal, “Adoption of cloud computing in e-government: A systematic literature review,” *Pertanika J. Sci. Technol.*, vol. 30, no. 1, pp. 655–689, 2022, doi: 10.47836/pjst.30.1.36.',
  ukeje2024: 'N. Ukeje, J. Gutierrez, and K. Petrova, “Information security and privacy challenges of cloud computing for government adoption: A systematic review,” *Int. J. Inf. Secur.*, vol. 23, pp. 1459–1475, 2024, doi: 10.1007/s10207-023-00797-6.',
  olnes2017: 'S. Ølnes, J. Ubacht, and M. Janssen, “Blockchain in government: Benefits and implications of distributed ledger technology for information sharing,” *Gov. Inf. Q.*, vol. 34, no. 3, pp. 355–364, 2017, doi: 10.1016/j.giq.2017.09.007.',
  tan2022: 'E. Tan, S. Mahula, and J. Crompvoets, “Blockchain governance in the public sector: A conceptual framework for public management,” *Gov. Inf. Q.*, vol. 39, no. 1, Art. no. 101625, 2022, doi: 10.1016/j.giq.2021.101625.',
  shahaab2023: 'A. Shahaab, I. A. Khan, R. Maude, C. Hewage, and Y. Wang, “Public service operational efficiency and blockchain – A case study of Companies House, UK,” *Gov. Inf. Q.*, vol. 40, no. 1, Art. no. 101759, 2023, doi: 10.1016/j.giq.2022.101759.',
  chatfield2019: 'A. T. Chatfield and C. G. Reddick, “A framework for Internet of Things-enabled smart government: A case of IoT cybersecurity policies and use cases in U.S. federal government,” *Gov. Inf. Q.*, vol. 36, no. 2, pp. 346–357, 2019, doi: 10.1016/j.giq.2018.09.007.',
  nist207: 'S. Rose, O. Borchert, S. Mitchell, and S. Connelly, “Zero trust architecture,” NIST, Gaithersburg, MD, USA, Special Publication 800-207, 2020, doi: 10.6028/NIST.SP.800-207.',
  nistcsf: 'National Institute of Standards and Technology, “The NIST Cybersecurity Framework (CSF) 2.0,” NIST CSWP 29, 2024, doi: 10.6028/NIST.CSWP.29.',
  iso27001: '*Information Security, Cybersecurity and Privacy Protection — Information Security Management Systems — Requirements*, ISO/IEC 27001:2022, 2022.',
  unesco2021: 'UNESCO, *Recommendation on the Ethics of Artificial Intelligence*. Paris, France: UNESCO, 2021.',
  unesco2025: 'UNESCO, *Data Governance Toolkit: Navigating Data in the Digital Age*. Paris, France: UNESCO, 2025.',
  peffers2007: 'K. Peffers, T. Tuunanen, M. A. Rothenberger, and S. Chatterjee, “A design science research methodology for information systems research,” *J. Manage. Inf. Syst.*, vol. 24, no. 3, pp. 45–77, 2007, doi: 10.2753/MIS0742-1222240302.',
  hevner2004: 'A. R. Hevner, S. T. March, J. Park, and S. Ram, “Design science in information systems research,” *MIS Q.*, vol. 28, no. 1, pp. 75–105, 2004.',
  venable2016: 'J. Venable, J. Pries-Heje, and R. Baskerville, “FEDS: A framework for evaluation in design science research,” *Eur. J. Inf. Syst.*, vol. 25, no. 1, pp. 77–89, 2016, doi: 10.1057/ejis.2014.36.',
  kazman2000: 'R. Kazman, M. Klein, and P. Clements, “ATAM: Method for architecture evaluation,” Software Eng. Inst., Carnegie Mellon Univ., Pittsburgh, PA, USA, Tech. Rep. CMU/SEI-2000-TR-004, 2000.',
  dictissp: 'Department of Information and Communications Technology, “Revised Information Systems Strategic Plan (ISSP) template guidelines,” DICT, Quezon City, Philippines, 2026.',
};

// ---------------------------------------------------------------- text helpers
// Inline markup: *italic*, **bold**, [@key] or [@a,@b] citations.
const order = [];
function cite(keys) {
  return keys.map(k => {
    if (!REFS[k]) throw new Error('Unknown ref ' + k);
    if (!order.includes(k)) order.push(k);
    return '[' + (order.indexOf(k) + 1) + ']';
  }).join(', ');
}
function resolveCites(s) {
  return s.replace(/\[@([^\]]+)\]/g, (_, g) => cite(g.split(',').map(x => x.trim().replace(/^@/, ''))));
}
function runs(text, base = {}) {
  text = resolveCites(text);
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, ...base }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ ...base, text: t.slice(2, -2), bold: true, font: FONT }));
    else out.push(new TextRun({ ...base, text: t.slice(1, -1), italics: true, font: FONT }));
    last = re.lastIndex;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, ...base }));
  return out;
}

const body = [];
const P = (t) => body.push(new Paragraph({
  children: runs(t, { size: 20 }), alignment: AlignmentType.JUSTIFIED,
  indent: { firstLine: 202 }, spacing: { before: 0, after: 0, line: 228, lineRule: 'exact' },
}));
const PL = (label, t) => body.push(new Paragraph({ // paragraph with italic run-in label
  children: [new TextRun({ text: label + ' ', italics: true, font: FONT, size: 20 }), ...runs(t, { size: 20 })],
  alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 202 },
  spacing: { before: 0, after: 0, line: 228, lineRule: 'exact' },
}));
let secNo = 0, subNo = 0;
const ROMAN = ['I', 'II', 'III', 'IV', 'V', 'VI'];
const H1 = (t) => { subNo = 0; body.push(new Paragraph({
  children: [new TextRun({ text: ROMAN[secNo++] + '. ' + t, smallCaps: true, font: FONT, size: 20 })],
  alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 }, keepNext: true,
})); };
const H2 = (t) => body.push(new Paragraph({
  children: [new TextRun({ text: String.fromCharCode(65 + subNo++) + '. ' + t, italics: true, font: FONT, size: 20 })],
  spacing: { before: 80, after: 40 }, keepNext: true,
}));

let figNo = 0;
function FIG(file, caption, wIn) {
  const buf = fs.readFileSync(path.join(__dirname, 'figures', file));
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const pxW = Math.round(wIn * 96);
  body.push(new Paragraph({
    children: [new ImageRun({ type: 'png', data: buf, transformation: { width: pxW, height: Math.round(pxW * h / w) } })],
    alignment: AlignmentType.CENTER, spacing: { before: 100, after: 40 }, keepNext: true,
  }));
  body.push(new Paragraph({
    children: [new TextRun({ text: `Fig. ${++figNo}. `, font: FONT, size: 16 }), ...runs(caption, { size: 16 })],
    alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 120 },
  }));
}

let tabNo = 0;
const TROMAN = ['I', 'II', 'III', 'IV', 'V'];
function TABLE(title, header, rows, widths) {
  body.push(new Paragraph({
    children: [new TextRun({ text: 'TABLE ' + TROMAN[tabNo++], font: FONT, size: 16 })],
    alignment: AlignmentType.CENTER, spacing: { before: 120, after: 0 }, keepNext: true,
  }));
  body.push(new Paragraph({
    children: [new TextRun({ text: title, smallCaps: true, font: FONT, size: 16 })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 }, keepNext: true,
  }));
  const line = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const mk = (cells, isHead, isLast) => new TableRow({
    cantSplit: true, tableHeader: isHead,
    children: cells.map((c, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      margins: { top: 15, bottom: 15, left: 50, right: 50 },
      verticalAlign: VerticalAlign.CENTER,
      shading: isHead ? { type: ShadingType.CLEAR, fill: 'E6E6E6', color: 'auto' } : undefined,
      borders: { top: isHead ? line : none, bottom: (isHead || isLast) ? line : { style: BorderStyle.SINGLE, size: 2, color: 'BFBFBF' }, left: none, right: none },
      children: [new Paragraph({
        children: runs(c, { size: 14, bold: isHead }),
        alignment: (i === 0 || isHead) ? AlignmentType.LEFT : AlignmentType.LEFT,
        spacing: { before: 0, after: 0, line: 170, lineRule: 'exact' },
      })],
    })),
  });
  body.push(new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [mk(header, true, false), ...rows.map((r, i) => mk(r, false, i === rows.length - 1))],
  }));
  body.push(new Paragraph({ children: [], spacing: { before: 0, after: 80 } }));
}

// ================================================================= CONTENT
H1('Introduction');
P('Digital transformation in government is not the conversion of paper into files; it is the redesign of processes, services, and inter-organizational relationships around shared data [@mergel2019]. Global benchmarks likewise treat integrated, whole-of-government service delivery—rather than isolated agency portals—as the mark of digital maturity [@un2024]. Philippine agencies, however, operate information systems that were procured at different times, for different mandates, and on different platforms. Records that describe a single public asset are therefore re-encoded, e-mailed, couriered, and reconciled manually across offices, multiplying effort and error.');
P('Two recent statutes make this problem urgent. Republic Act (RA) 12254, the E-Governance Act, mandates the digitalization of government processes and the interoperability of government information systems under a whole-of-government approach [@ra12254]. RA 12234, the Konektadong Pinoy Act, liberalizes entry into the data-transmission industry, promotes infrastructure sharing, and opens the market to new providers, including satellite operators, to widen affordable connectivity [@ra12234]. Their implementing rules took effect in November 2025 and May 2026, respectively. The first law creates an obligation to interoperate; the second governs the networks on which that interoperability depends. Yet the E-Governance mandate presumes connectivity that the Konektadong Pinoy reforms are still building—a mismatch that is fundamentally a Network Enterprise Architecture (NEA) problem.');
P('The verification of public school-site land titles illustrates the problem concretely. To regularize a school site, a Department of Education (DepEd) Schools Division Office (SDO) must assemble the Original or Transfer Certificate of Title (OCT/TCT) from the Registry of Deeds under the Land Registration Authority (LRA), the tax declaration from the local assessor, the approved survey plan from the Department of Environment and Natural Resources (DENR), and deeds of donation from local government units (LGUs) or donors. In the author’s practitioner experience in an SDO school-site titling office, these artifacts arrive through letters, visits, and scanned uploads to a document management system, cloud drives, and spreadsheets—each step a re-encoding, none governed by a common identity, data standard, or security policy. The same pattern recurs in permits, social protection, and disaster response.');
P('**Problem statement.** Without an architecture that jointly addresses connectivity, security, data, and governance, mandated interoperability risks producing either point-to-point integrations whose complexity grows quadratically, or centralized repositories that concentrate risk and strain the proportionality principle of the Data Privacy Act [@ra10173]. This study therefore asks: **RQ1**—What NEA requirements arise from the literature, Philippine law, and local operating conditions for cross-agency records interoperability? **RQ2**—What architecture satisfies these requirements? **RQ3**—How does it compare with siloed and centralized alternatives on technical criteria?');
P('The contributions are: (1) eight literature- and law-traceable architectural requirements; (2) PH-NEA, a reference architecture of five layers and two cross-cutting planes featuring a federated zero-trust interoperability gateway and a multipath, offline-tolerant access network; (3) a principle-to-control mapping that operationalizes data-privacy and AI-ethics principles; and (4) an ex ante evaluation using quality-attribute scenarios, analytical models, and comparative scoring. Section II reviews the literature, Section III describes the method, Section IV presents and evaluates the architecture, and Section V concludes.');

H1('Literature Review');
P('The review condenses an earlier critical literature review into five themes. Table I summarizes each theme’s consensus, its limitation for the Philippine setting, and the requirement it yields.');
H2('Digital Transformation and Enterprise Architecture');
P('Enterprise architecture (EA) is consistently framed as the instrument that aligns processes, data, applications, and technology during public-sector transformation [@gong2020]. Janssen [@janssen2012] adds a critical qualification: interoperability across agencies is sociopolitical, negotiated among autonomous bodies rather than imposed by technology. Success factors identified in government EA—strategic alignment, governance, skills, and stakeholder engagement—vary widely by agency maturity [@juraida2024]. Frameworks such as TOGAF [@togaf] and ArchiMate [@archimate] provide method and notation but implicitly assume a single enterprise under one authority. A whole-of-government architecture is instead a federation of autonomous agencies at different maturity levels, which the generic frameworks do not resolve.');
H2('Interoperability and Data Governance');
P('Comparative work shows that successful e-government programs pair an interoperability framework with EA [@guijarro2007]; the European Interoperability Framework separates legal, organizational, semantic, and technical interoperability [@eif2017]. Two integration models compete: central consolidation into a shared repository, and federated exchange in which data stay with authoritative sources and are retrieved on demand. Data-governance studies favor the latter because accountability follows data stewardship [@janssen2020], [@vandonge2022], and Estonia’s X-Road demonstrates federated exchange through signed, logged messages between autonomous registries [@xroad]. The limitation is contextual: these models assume the dependable, high-bandwidth networks of European states, which many Philippine LGUs lack.');
H2('Emerging Technologies as Conditional Enablers');
P('Studies of artificial intelligence (AI) [@sousa2019], [@straub2023], the Internet of Things (IoT) [@kankanhalli2019], [@wirtz2019], cloud [@abied2022], [@ukeje2024], and blockchain [@olnes2017], [@tan2022], [@shahaab2023] converge on one conclusion despite examining technologies in isolation: each delivers value only when the underlying data, network, and governance are sound. AI requires quality data and human oversight; cloud adoption is constrained chiefly by security and privacy concerns; and blockchain is justified only when several parties need a tamper-evident shared record without a trusted operator. The synthesis implies that technologies should enter an architecture through explicit selection criteria, not by trend.');
H2('Security, Privacy, and Ethics');
P('Interconnection enlarges the attack surface [@chatfield2019], and perimeter defenses break down when traffic crosses agency boundaries. Zero-trust architecture (ZTA) replaces implicit network trust with per-request authentication and authorization [@nist207], while the NIST CSF 2.0 [@nistcsf] and ISO/IEC 27001 [@iso27001] supply governance and management-system controls. UNESCO’s AI-ethics recommendation and data-governance toolkit stress proportionality, transparency, and human oversight [@unesco2021], [@unesco2025], echoing the Data Privacy Act’s principles of legitimate purpose and proportionality [@ra10173]. However, these principles are seldom translated into concrete, verifiable architectural controls.');
H2('Connectivity and the Research Gap');
P('Finally, e-government studies treat connectivity as a given, while connectivity studies stop at access and rarely reach the application layer. In an archipelago exposed to typhoons and earthquakes, with uneven LGU broadband, cloud services and application programming interfaces (APIs) fail without resilient access. **Research gap:** to the author’s knowledge, no study offers a Philippine reference NEA that jointly (a) links the E-Governance Act’s interoperability mandate with the Konektadong Pinoy Act’s connectivity provisions, (b) tolerates intermittent connectivity and uneven agency readiness, and (c) traces legal and ethical principles to network and security controls.');

TABLE('Literature Synthesis and Derived Requirements',
  ['Theme', 'Consensus', 'Limitation (PH)', 'Req.'],
  [
    ['EA & transformation [@mergel2019], [@gong2020], [@juraida2024]', 'EA aligns process, data, technology; change is organizational', 'Assumes one authority, uniform maturity', 'R1'],
    ['Interoperability & data [@eif2017], [@janssen2020], [@vandonge2022]', 'Federated exchange with stewardship; shared semantics', 'Assumes reliable high-bandwidth networks', 'R2, R3'],
    ['Emerging tech [@straub2023], [@ukeje2024], [@tan2022]', 'Value conditional on data, network, governance', 'Studied in isolation; trend-driven', 'R4'],
    ['Security & ethics [@nist207], [@unesco2021], [@ra10173]', 'Zero trust; privacy and oversight by design', 'Principles not mapped to controls', 'R5, R6'],
    ['Connectivity & policy [@ra12234], [@ra12254], [@un2024]', 'Connectivity is a precondition of e-services', 'Rarely modeled inside EA; hazard exposure', 'R7, R8'],
  ], [1320, 1500, 1580, 640]);

H1('Methodology');
H2('Research Design');
P('The study follows the Design Science Research Methodology (DSRM) [@peffers2007] because its aim is a purposeful artifact—an architecture—whose utility must be demonstrated rather than a hypothesis tested; it observes the design-science guidelines of relevance, rigor, and evaluation [@hevner2004]. Fig. 1 maps the six DSRM activities to this study. TOGAF Architecture Development Method (ADM) Phases A–D [@togaf] serve as the design procedure within Activity 3, and the models use ArchiMate-style layered viewpoints [@archimate].');
FIG('fig1_method.png', 'Research method: DSRM activities with TOGAF ADM Phases A–D embedded in design and a four-part ex ante evaluation.', 3.45);
H2('Data Sources');
P('Three sources informed the problem and requirements. (i) *Literature:* peer-reviewed journal articles (mainly 2017–2026, plus seminal works), international standards, and EA frameworks, screened for relevance to digital government, EA, interoperability, emerging technologies, security, or Philippine digital transformation, then coded into the five themes of Table I. (ii) *Legislation:* RA 12234, RA 12254, RA 10173, and RA 11055 [@ra11055], analyzed for architectural obligations. (iii) *Practitioner case:* the author’s work in school-site titling at a DepEd SDO and in records management at a DENR regional office, used as a reflective-practitioner source to model the As-Is process. No personal or record-level data were used.');
H2('Design Procedure');
P('Phase A (Vision) scoped the architecture to cross-agency records exchange and identified stakeholders: DepEd, LGU assessors, LRA, DENR, the Department of Information and Communications Technology (DICT), the National Privacy Commission, and the Philippine Statistics Authority as PhilSys operator. Phase B (Business) modeled the As-Is and To-Be verification processes. Phase C (Data/Application) defined a canonical School-Site Record entity and the application services. Phase D (Technology) specified the network topology, security, and recovery design. Each design decision was recorded with its rationale and traced to requirements R1–R8.');
H2('Evaluation Design');
P('Because the artifact is not yet implemented, the evaluation follows the *ex ante, artificial* strategy of the FEDS framework [@venable2016], which suits early-stage, high-risk designs. Four techniques were used: (E1) requirement traceability; (E2) quality-attribute scenarios adapted from the Architecture Tradeoff Analysis Method (ATAM) [@kazman2000], comparing As-Is and To-Be responses; (E3) analytical models—availability of parallel independent links, A = 1 − Π(1 − A*i*), and integration complexity, N(N − 1)/2 point-to-point interfaces versus N hub interfaces; and (E4) a comparative matrix of three alternatives over ten criteria, scored 1–5 against an anchored rubric (1 = unaddressed, 3 = partially addressed with known risk, 5 = addressed by an explicit mechanism) with literature evidence for each score.');

H1('Results and Discussion');
H2('Current-State (As-Is) Analysis');
P('The As-Is verification process has seven steps: the school submits documents; the SDO scans and encodes them; the SDO requests the title status from the Registry of Deeds, the tax declaration from the assessor, and the survey status from DENR, each by letter or visit; staff cross-check the documents manually; and the results are re-encoded into spreadsheets for reporting. The model shows six inter-office handoffs and three re-encodings of the same lot data. The analysis exposes limitations in every architectural domain. *Connectivity:* municipal offices typically depend on a single consumer-grade broadband link, so an outage halts online work. *Interoperability:* no APIs exist, and documents travel as scanned PDFs. *Data:* identifiers (school ID, lot number, title number, tax-declaration number) have no crosswalk. *Security:* shared folder links and personal accounts substitute for identity management, and no log records who viewed a title. *Governance:* exchanges rest on memoranda rather than data-sharing agreements. *Infrastructure:* records reside on local or consumer storage without disaster recovery (DR).');
H2('Proposed Network Enterprise Architecture (PH-NEA)');
P('Fig. 2 presents PH-NEA. A key refinement over the earlier seven-layer proposal is that security and governance are modeled as vertical planes rather than layers, because both must constrain every layer—a “security layer” sitting between other layers misrepresents how zero trust operates.');
FIG('fig2_architecture.png', 'PH-NEA reference architecture: five service layers and two cross-cutting planes. L3 (bold) is the integration core.', 3.45);
PL('L1 Connectivity and access.', 'Each office receives a software-defined WAN (SD-WAN) edge that bonds a primary wired link and a secondary wireless or low-Earth-orbit (LEO) satellite link from a *different* provider and medium—an option the Konektadong Pinoy Act’s open, multi-provider market makes procurable. Encrypted overlays carry traffic over GovNet or public transport, and quality-of-service (QoS) policies prioritize interoperability traffic. A store-and-forward queue at the edge holds outbound requests during outages, making the exchange *offline-tolerant* rather than offline-failing (R7).');
PL('L2 Shared infrastructure.', 'Agency services run in GovCloud or the national government data centers, with asynchronous replication to an out-of-region DR site so that a typhoon or earthquake affecting one region does not take down both (R8). The targets are a recovery-point objective (RPO) of 15 minutes or less and a recovery-time objective (RTO) of 4 hours or less.');
PL('L3 Data and interoperability.', 'A federated Government Interoperability Gateway follows the X-Road pattern [@xroad]: each source agency runs a security server, data remain at the authoritative source (R2), and requests are mutually authenticated, signed, and logged. API contracts, a service catalog, a canonical School-Site Record, and an identifier crosswalk provide shared semantics; validation rules flag discrepancies for a named data steward (R3).');
PL('L4 Applications and emerging technologies.', 'Technologies pass explicit decision gates (R4). AI-based field extraction from scanned titles is admitted because the task is repetitive and verifiable, but extractions below a confidence threshold go to human review. Blockchain is *not* used for the general record store: signed gateway logs already provide tamper evidence under a trusted operator. Hash anchoring is reserved for multi-agency title-status events, if a later governance review finds no acceptable trusted operator.');
PL('L5 Services.', 'A composite “verify school site” service replaces the six handoffs; citizens and donors access related services through the eGovPH channel, with assisted channels for users who have limited connectivity.');
PL('Plane A, zero-trust security and identity.', 'Policy decision and enforcement points authorize every request using identity, device posture, and *purpose* attributes [@nist207]; staff authenticate with multi-factor authentication (MFA), citizens through PhilSys [@ra11055], and agencies through PKI-based mutual TLS. Micro-segmentation and a security operations center (SOC) under an ISO/IEC 27001 management system complete the plane (R5).');
PL('Plane B, governance.', 'An EA board under DICT maintains standards, decision gates, and data-sharing agreements reviewed for Data Privacy Act compliance. Readiness tiers implement R1: Tier 0 offices (manual records) are served through an assisted regional portal; Tier 1 offices consume services through a web portal; Tier 2 offices publish APIs through a hosted security server; and Tier 3 offices run their own security server, AI, and analytics. Agencies can join at their current tier and advance incrementally.');
H2('Architectural Models: Network Topology and Data Flow');
P('Fig. 3 instantiates PH-NEA for school-site verification. A titling officer authenticates with MFA (1); the policy decision point issues a purpose-bound scope that limits the response to title status, lot, and area rather than the full title image (2); the gateway sends parallel requests to the Registry of Deeds, assessor, and DENR security servers (3–4); the responses are validated against the canonical model, discrepancies are routed to a steward, and every exchange is logged (5); and a signed, consolidated record returns to the document management system (6). The LGU assessor, at Tier 1–2, uses a hosted security server—showing that a low-capacity office can still participate.');
FIG('fig3_topology.png', 'Logical network topology and To-Be data flow for school-site verification. Solid lines show primary fiber; dashed lines show diverse backup links.', 3.45);
P('Table II operationalizes requirement R6 by tracing each legal or ethical principle to a verifiable control and its location in the architecture.');
TABLE('Principle-to-Control Traceability (R6)',
  ['Principle (source)', 'PH-NEA control', 'Loc.'],
  [
    ['Proportionality, legitimate purpose [@ra10173]', 'Purpose-bound API scopes; attribute-level responses, not full documents', 'L3, A'],
    ['Transparency, accountability [@unesco2025]', 'Signed, immutable exchange log; access history viewable by data subject', 'L3, L5'],
    ['Human oversight [@unesco2021]', 'Confidence threshold routes AI extractions to human review', 'L4'],
    ['Security of processing [@iso27001]', 'MFA, mTLS, per-request PDP/PEP, SOC monitoring', 'A'],
    ['Inclusion [@ra12234], [@unesco2021]', 'Assisted channels; offline-tolerant edge; Tier 0 entry', 'L1, L5'],
    ['Data quality, integrity [@janssen2020]', 'Canonical-model validation; steward workflow; ID crosswalk', 'L3, B'],
  ], [1600, 2860, 580]);

H2('Evaluation and Validation');
PL('E1 Traceability.', 'Every requirement maps to at least one explicit mechanism: R1 to the readiness tiers; R2 to the federated security servers; R3 to the canonical model and crosswalk; R4 to the decision gates; R5 to Plane A; R6 to Table II; R7 to the SD-WAN and edge queue; and R8 to out-of-region DR.');
PL('E2 Quality-attribute scenarios.', 'Table III summarizes six scenarios. Four are met by explicit mechanisms. Two are met only partially, because their response measures depend on DICT data-center capacity and on load testing that only implementation can provide.');
TABLE('Quality-Attribute Scenario Evaluation (ATAM-Inspired)',
  ['Scenario (stimulus)', 'As-Is response', 'PH-NEA response', 'Result'],
  [
    ['S1 Avail.: assessor fiber cut in typhoon', 'Requests stall; courier fallback', 'SD-WAN failover to LEO; edge queue retains requests', 'Met'],
    ['S2 Security: SDO credential phished', 'Shared links expose all titles', 'MFA + device posture; per-request scope limits blast radius; SOC alert', 'Met'],
    ['S3 Interop.: new agency joins', 'New bespoke integrations', 'Catalog entry, canonical schema, one certificate', 'Met'],
    ['S4 Data: lot no. differs across title and tax declaration', 'Found late, at reporting', 'Gateway validation flags it; steward resolves', 'Met'],
    ['S5 DR: regional data center lost', 'Ad hoc local copies', 'Out-of-region replica; RPO 15 min, RTO 4 h', 'Partial'],
    ['S6 Scale: national titling drive', 'Manual backlog', 'Rate limits, autoscaling', 'Partial'],
  ], [1300, 1060, 2040, 640]);
PL('E3 Analytical models.', 'Assume, illustratively, that a municipal office’s single broadband link has 99.0% availability (87.6 h of downtime per year). Adding an independent link of equal availability over a different medium and provider gives 1 − (0.01)² = 99.99%, or about 0.9 h per year. The independence assumption fails if both links share site power, which is why Fig. 3 includes UPS/solar backup, and the edge queue preserves transactions during any residual downtime. For integration complexity, connecting N agencies point-to-point requires N(N − 1)/2 interfaces versus N for the gateway: 10 versus 5 for the five offices in the case, and 190 versus 20 for a 20-agency deployment. In the process model, inter-office handoffs fall from six to one (a single steward review), and re-encodings fall from three to zero.');
PL('E4 Comparative evaluation.', 'Table IV compares the siloed status quo (A0), a centralized national records database (A1), and PH-NEA (A2). PH-NEA scores highest (38 of 50), but it is not dominant: the centralized option is easier to manage, and both alternatives cost about the same or less.');
TABLE('Comparative Evaluation of Alternatives (1–5)',
  ['Criterion', 'A0 Siloed', 'A1 Central', 'A2 PH-NEA'],
  [
    ['Interoperability', '1', '4', '5'],
    ['Security', '2', '3', '4'],
    ['Availability / resilience', '2', '3', '4'],
    ['Scalability', '2', '4', '4'],
    ['Privacy compliance (RA 10173)', '2', '2', '4'],
    ['Manageability', '2', '4', '3'],
    ['Cost-effectiveness', '2', '3', '3'],
    ['Implementation feasibility', '5', '2', '3'],
    ['Inclusion (low connectivity)', '2', '2', '4'],
    ['Agency autonomy / accountability', '5', '1', '4'],
    ['**Total (of 50)**', '**25**', '**28**', '**38**'],
  ], [2280, 920, 920, 920]);

H2('Discussion');
P('*Major findings.* The results indicate that cross-agency interoperability in the Philippines is as much a network and governance problem as an application problem. The decisive design choice is federation: A1 matches PH-NEA on interoperability and scalability but loses on privacy, autonomy, and inclusion because it concentrates sensitive land and identity data in a single high-value target and strips agencies of stewardship—supporting Janssen’s sociopolitical view of interoperability [@janssen2012] and the stewardship literature [@vandonge2022]. Making connectivity a first-class layer converts the Konektadong Pinoy Act from background policy into a design lever: diverse multi-provider access is what makes the modeled 99.99% availability achievable in municipalities.');
P('*Trade-offs and feasibility.* PH-NEA demands more operational capability—security servers, PKI life-cycle management, a SOC, and an active EA board—which explains its lower manageability score. Agencies differ in skills and funding [@juraida2024], so the readiness tiers matter: a Tier 0 or Tier 1 LGU can participate through hosted components while capacity grows, and the investments can be phased through agency Information Systems Strategic Plans [@dictissp].');
P('*Legislative alignment and implications.* PH-NEA operationalizes the E-Governance Act’s interoperability mandate through L3 and Plane B; it uses the Konektadong Pinoy Act’s multi-provider market in L1; it embeds Data Privacy Act proportionality in purpose-bound scopes; and it reuses PhilSys for citizen identity. For Philippine NEA practice, the implication is a shift in procurement from standalone systems to published API contracts and service-level commitments that every new system must meet.');
P('*Limitations.* The evaluation is ex ante and analytical. The scores come from a single evaluator, although they are anchored to an explicit rubric and literature evidence. The availability figures are assumptions, not measurements. The demonstration covers one process, and no cost model was developed.');

H1('Conclusion');
P('This study set out to determine how a Network Enterprise Architecture can enable secure, resilient cross-agency records interoperability under the E-Governance and Konektadong Pinoy Acts. Using DSRM with TOGAF ADM, it derived eight requirements from the literature, legislation, and a practitioner case, and designed PH-NEA: five layers and two cross-cutting planes centered on a federated zero-trust gateway and an offline-tolerant multipath access network. Ex ante evaluation showed that all eight requirements are traceable to explicit mechanisms and that four of six quality scenarios are fully met. In the analytical models, link availability rises from 99.0% to 99.99% and verification handoffs fall from six to one. PH-NEA also outscores siloed and centralized alternatives, at the cost of higher governance overhead.');
P('The principal contribution is a legislation-traceable reference architecture that treats connectivity, zero trust, and agency readiness as integral parts of interoperability design rather than as afterthoughts. For Philippine digital transformation, it offers a phased path through which low-capacity LGUs and mature agencies can join one exchange fabric. Future work should validate PH-NEA with DICT, DepEd, and LGU architects through a Delphi study, emulate SD-WAN failover and gateway load to measure the scenario responses, develop a cost model, and pilot the design in one SDO with two LGUs.');

// ---------------------------------------------------------------- references section
const refHead = new Paragraph({
  children: [new TextRun({ text: 'References', smallCaps: true, font: FONT, size: 20 })],
  alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 }, keepNext: true,
});
const refParas = order.map((k, i) => new Paragraph({
  children: [new TextRun({ text: `[${i + 1}]\t`, font: FONT, size: 16 }), ...runs(REFS[k], { size: 16 })],
  tabStops: [{ type: 'left', position: 360 }],
  indent: { left: 360, hanging: 360 }, alignment: AlignmentType.JUSTIFIED,
  spacing: { before: 0, after: 10, line: 190, lineRule: 'exact' },
}));
const unused = Object.keys(REFS).filter(k => !order.includes(k));
if (unused.length) console.warn('Unused references:', unused.join(', '));

// ---------------------------------------------------------------- front matter
const center = (children, after = 0, before = 0) => new Paragraph({ children, alignment: AlignmentType.CENTER, spacing: { before, after } });
const front = [
  center([new TextRun({ text: 'A Connectivity-Aware, Zero-Trust Network Enterprise Architecture for Cross-Agency Records Interoperability in Philippine Government: A Design Science Study', font: FONT, size: 48 })], 240),
  center([new TextRun({ text: 'Jeanalyn Marie B. Lagera', font: FONT, size: 22 })]),
  center([new TextRun({ text: 'MIT-O, Network Enterprise Architecture', italics: true, font: FONT, size: 20 })]),
  center([new TextRun({ text: '[University Name]', italics: true, font: FONT, size: 20 })]),
  center([new TextRun({ text: '[City], Philippines', font: FONT, size: 20 })]),
  center([new TextRun({ text: '[email address]', font: FONT, size: 20 })], 240),
];
const abstractText = 'Philippine government agencies operate independently procured information systems, so records describing a single public asset are re-encoded, couriered, and reconciled manually across offices. Republic Act (RA) 12254, the E-Governance Act, mandates interoperable government information systems, while RA 12234, the Konektadong Pinoy Act, reforms the data-transmission infrastructure on which interoperability depends; the literature, however, rarely links the two at the architectural level. This paper investigates how a Network Enterprise Architecture can enable secure and resilient cross-agency records exchange under uneven connectivity and agency readiness. Following the Design Science Research Methodology, with TOGAF ADM Phases A–D as the design procedure, eight requirements were derived from a condensed critical literature review, Philippine legislation, and a practitioner-grounded case: the verification of public school-site land titles across the Department of Education, local assessors, the Registry of Deeds, and the Department of Environment and Natural Resources. The resulting artifact, PH-NEA, comprises five layers and two cross-cutting planes—zero-trust security and governance—built around a federated interoperability gateway, a multipath offline-tolerant access network, principle-to-control traceability, and agency readiness tiers. Ex ante evaluation with six quality-attribute scenarios, analytical models, and a ten-criterion comparison against siloed and centralized alternatives shows that PH-NEA addresses all eight requirements. In the analytical models, it raises modeled office availability from 99.0% to 99.99% and reduces verification handoffs from six to one, at the cost of higher governance overhead. The study contributes a reusable, legislation-traceable reference architecture for Philippine digital government.';
const abs = new Paragraph({
  children: [new TextRun({ text: 'Abstract', bold: true, italics: true, font: FONT, size: 18 }),
    new TextRun({ text: '—' + abstractText, bold: true, font: FONT, size: 18 })],
  alignment: AlignmentType.JUSTIFIED, spacing: { after: 120 },
});
const kw = new Paragraph({
  children: [new TextRun({ text: 'Keywords', bold: true, italics: true, font: FONT, size: 18 }),
    new TextRun({ text: '—network enterprise architecture, interoperability, zero-trust architecture, design science research, Philippine e-governance', bold: true, font: FONT, size: 18 })],
  alignment: AlignmentType.JUSTIFIED, spacing: { after: 120 },
});

const page = { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1440, left: 900, right: 900 } };
const doc = new Document({
  creator: 'Jeanalyn Marie B. Lagera',
  title: 'CS Deliverable 2 – IEEE Conference Paper',
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  sections: [
    { properties: { page }, children: front },
    { properties: { page, type: SectionType.CONTINUOUS, column: { count: 2, space: 360, equalWidth: true } },
      children: [abs, kw, ...body, refHead, ...refParas] },
  ],
});
const words = abstractText.split(/\s+/).length;
console.log('Abstract words:', words, '| references:', order.length);
Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(path.join(__dirname, 'CS2_Lagera_Jeanalyn_Marie_B_IEEE.docx'), b);
  console.log('written');
});
