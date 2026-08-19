import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "C:/Users/woody/Workspace/vault/documents/presentations/demo-v3";
const OUT = `${ROOT}/when-assistant-can-act-v3.pptx`;
const QA = "C:/Users/woody/Documents/Codex/2026-08-19/presentations-plugin-presentations-openai-primary-runtime-2/work/demo-v3-build/qa";
const ASSET = `${ROOT}/assets`;
const C = { bg:"#07090D", panel:"#10141C", panel2:"#171C26", line:"#29313E", ink:"#F7F4EC", muted:"#A8ADB7", violet:"#A58BFF", cyan:"#48E5D1", amber:"#FFBD63", red:"#FF6477", white:"#FFFFFF" };
const SANS = "Aptos Display";
const MONO = "Aptos Mono";
const pres = Presentation.create({ slideSize:{ width:1280, height:720 } });

function shp(slide, geometry, position, fill="none", lineFill="none", width=0, name){
  return slide.shapes.add({geometry, name, position, fill, line:{style:"solid",fill:lineFill,width}});
}
function txt(slide, value, pos, o={}){
  const s=shp(slide,"textbox",pos,o.fill||"none",o.line||"none",o.lineWidth||0,o.name);
  s.text=value;
  s.text.style={fontSize:o.size||24,bold:!!o.bold,color:o.color||C.ink,typeface:o.mono?MONO:SANS,alignment:o.align||"left",verticalAlignment:o.vAlign||"top",lineSpacing:o.lineSpacing||1,autoFit:"shrinkText",insets:{left:0,right:0,top:0,bottom:0}};
  return s;
}
function rich(slide,runs,pos,o={}){
  const s=shp(slide,"textbox",pos,"none","none",0,o.name);
  s.text.style={fontSize:o.size||54,bold:!!o.bold,color:o.color||C.ink,typeface:o.mono?MONO:SANS,lineSpacing:o.lineSpacing||.96,autoFit:"shrinkText",insets:{left:0,right:0,top:0,bottom:0}};
  s.text.set([runs.map(r=>typeof r==="string"?r:{run:r.run,textStyle:{bold:r.bold,color:r.color||o.color||C.ink,fontSize:`${r.size||o.size||54}px`,typeface:r.mono?MONO:SANS}})]);
  return s;
}
function rule(slide,x,y,w,color=C.line,h=1){shp(slide,"rect",{left:x,top:y,width:w,height:h},color,"none",0)}
function chrome(slide,chapter,page,total=12,right="AI PRIMITIVES · TOOLS"){
  slide.background.fill=C.bg;
  txt(slide,chapter.toUpperCase(),{left:64,top:34,width:520,height:20},{size:12,bold:true,color:C.cyan,mono:true});
  txt(slide,right,{left:760,top:34,width:456,height:20},{size:12,bold:true,color:C.muted,mono:true,align:"right"});
  txt(slide,`${String(page).padStart(2,"0")} / ${String(total).padStart(2,"0")}`,{left:1100,top:680,width:116,height:14},{size:10,color:C.muted,mono:true,align:"right"});
}
function note(slide,body,sources=[]){
  const block=sources.length?`\n\n[Sources]\n${sources.map(s=>`- ${s}`).join("\n")}`:"";
  slide.speakerNotes.textFrame.setText(body+block);
}
async function imageBlob(file){const b=await fs.readFile(file);return b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength)}
async function addImage(slide,file,pos,alt,fit="cover"){
  slide.images.add({blob:await imageBlob(file),contentType:"image/png",alt,fit,position:pos});
}
function kicker(slide,value,x=64,y=105,color=C.cyan){txt(slide,value.toUpperCase(),{left:x,top:y,width:520,height:22},{size:13,bold:true,color,mono:true})}

// 1 — cinematic opening
{
  const s=pres.slides.add(); s.background.fill=C.bg;
  await addImage(s,`${ASSET}/hero-control-gate.png`,{left:0,top:0,width:1280,height:720},"Abstract violet intelligence passing through an amber authorization gate into cyan action");
  shp(s,"rect",{left:0,top:0,width:560,height:720},"#07090D","none",0);
  txt(s,"AI PRIMITIVES · 04",{left:64,top:46,width:360,height:20},{size:12,bold:true,color:C.cyan,mono:true});
  kicker(s,"The moment software becomes an actor",64,150);
  rich(s,[{run:"When the assistant\ncan ",bold:true},{run:"act.",bold:true,color:C.violet}],{left:64,top:205,width:610,height:220},{size:80,lineSpacing:.9});
  txt(s,"The model proposes. The harness decides.\nTools touch the world.",{left:68,top:470,width:470,height:78},{size:25,color:"#D1D4DA",lineSpacing:1.12});
  note(s,"Talking points:\n• Open cold: ‘What changed the first time an assistant asked permission to run something?’\n• The model itself did not suddenly gain hands. The product around it gained a path into the world.\n• Use the visual left to right: violet is the model’s proposal, amber is authorization, cyan is action.\n• Central idea for the session: agency is assembled around the model.\n• Transition: before we follow that path, separate instructions from capabilities.",["AI-generated visual: OpenAI image generation, created for this deck"]);
}

// 2 — the distinction
{
  const s=pres.slides.add(); chrome(s,"01 · the distinction",2,12,"INSTRUCTION ≠ CAPABILITY");
  kicker(s,"One line changes everything");
  txt(s,"A skill explains the job.",{left:64,top:170,width:880,height:72},{size:58,bold:true});
  txt(s,"A tool changes the state.",{left:64,top:265,width:900,height:72},{size:58,bold:true,color:C.cyan});
  rule(s,64,390,1152,C.line,1);
  txt(s,"SKILL",{left:64,top:425,width:120,height:20},{size:13,bold:true,color:C.violet,mono:true});
  txt(s,"“Review a pull request this way.”",{left:205,top:412,width:760,height:42},{size:28});
  txt(s,"TOOL",{left:64,top:495,width:120,height:20},{size:13,bold:true,color:C.cyan,mono:true});
  txt(s,"“Approve pull request 42.”",{left:205,top:482,width:760,height:42},{size:28});
  txt(s,"HARNESS",{left:64,top:565,width:120,height:20},{size:13,bold:true,color:C.amber,mono:true});
  txt(s,"Decides whether that action runs.",{left:205,top:552,width:800,height:42},{size:28});
  note(s,"Talking points:\n• A skill is guidance: it teaches a repeatable way to perform a job.\n• A tool is capability: it exposes an operation that can read or change something.\n• The harness is the control plane: it decides what context is supplied, which tools are available, and whether a request runs.\n• Stress the verbs: explains, changes, decides.\n• Transition: once a tool is available, every action follows the same basic path.");
}

// 3 — control plane
{
  const s=pres.slides.add(); chrome(s,"02 · the control plane",3,12,"AGENCY IS ASSEMBLED");
  kicker(s,"The model never runs a command");
  txt(s,"Every action crosses a boundary.",{left:64,top:145,width:1000,height:72},{size:54,bold:true});
  const stages=[
    ["READ","Context enters",C.violet],
    ["PROPOSE","A tool request forms",C.violet],
    ["AUTHORIZE","Policy or a person decides",C.amber],
    ["ACT","The tool runs",C.cyan],
    ["PROVE","Evidence returns",C.cyan]
  ];
  stages.forEach((d,i)=>{const x=64+i*230; rule(s,x,305,190,d[2],4); txt(s,String(i+1).padStart(2,"0"),{left:x,top:330,width:50,height:18},{size:11,bold:true,color:d[2],mono:true}); txt(s,d[0],{left:x,top:375,width:195,height:40},{size:25,bold:true,color:d[2]}); txt(s,d[1],{left:x,top:438,width:190,height:54},{size:17,color:C.muted}); if(i<4) txt(s,"→",{left:x+194,top:375,width:36,height:36},{size:26,color:C.line,align:"center"});});
  txt(s,"The safety seam is AUTHORIZE.",{left:64,top:560,width:760,height:54},{size:34,bold:true,color:"#D8D9DB"});
  note(s,"Talking points:\n• Walk the audience from left to right: read, propose, authorize, act, prove.\n• The model produces a structured request; it does not execute the operation itself.\n• Authorization may be automatic policy or a visible human approval.\n• The result must come back as evidence, otherwise the loop cannot know what happened.\n• Point at AUTHORIZE: this is the safety seam.\n• Transition: this pattern becomes valuable when one contract reaches many systems.");
}

// 4 — visual enterprise contract
{
  const s=pres.slides.add(); s.background.fill=C.bg;
  await addImage(s,`${ASSET}/one-contract.png`,{left:0,top:0,width:1280,height:720},"Many enterprise system signals converging through one violet prism into a stable output");
  shp(s,"rect",{left:0,top:0,width:520,height:720},"#07090D","none",0);
  txt(s,"03 · THE ENTERPRISE GAP",{left:64,top:42,width:400,height:20},{size:12,bold:true,color:C.cyan,mono:true});
  kicker(s,"Paperwork Nightmare CLI",64,145);
  txt(s,"Twenty-two systems.\nOne contract.",{left:64,top:200,width:520,height:158},{size:60,bold:true,lineSpacing:.9});
  txt(s,"pncli turns the systems developers already use into structured JSON an assistant can reason over.",{left:68,top:420,width:430,height:112},{size:23,color:"#D1D4DA",lineSpacing:1.14});
  note(s,"Talking points:\n• pncli is deterministic TypeScript, not an AI model.\n• It provides one predictable interface across systems developers already use.\n• That reduces integration friction for the assistant and makes behavior easier to test.\n• The prism is the metaphor: many different inputs become one stable contract.\n• Avoid presenting ‘22’ as magic; the architectural point is consistency.\n• Transition: the contract matters because its output can become the next reliable decision.",["C:/Users/woody/Workspace/vault/wiki/pncli/pncli-overview.md","AI-generated visual: OpenAI image generation, created for this deck"]);
}

// 5 — structured output
{
  const s=pres.slides.add(); chrome(s,"04 · structured in, structured out",5,12,"JSON BECOMES THE NEXT DECISION");
  kicker(s,"Why the contract matters");
  rich(s,[{run:"Not a screenshot.\nNot scraped prose.\n",bold:true},{run:"A result you can test.",bold:true,color:C.violet}],{left:64,top:165,width:560,height:250},{size:50,lineSpacing:.94});
  shp(s,"roundRect",{left:700,top:125,width:516,height:440},C.panel,"#394455",1);
  shp(s,"ellipse",{left:728,top:151,width:11,height:11},C.red); shp(s,"ellipse",{left:750,top:151,width:11,height:11},C.amber); shp(s,"ellipse",{left:772,top:151,width:11,height:11},C.cyan);
  txt(s,"› pncli jira get-issue --key APP-901\n\n{\n  \"status\": \"To Do\",\n  \"labels\": [\"needs-repro\"],\n  \"url\": \"https://…/checkout\"\n}\n\n# inspect → compare → decide",{left:735,top:205,width:430,height:300},{size:18,color:C.cyan,mono:true,lineSpacing:1.22});
  txt(s,"Predictable contracts create reliable loops.",{left:64,top:520,width:560,height:60},{size:28,bold:true,color:"#D4D6DA"});
  note(s,"Talking points:\n• A screenshot is evidence for a person, but usually a poor contract for a workflow.\n• Scraped prose is fragile because structure and meaning are implicit.\n• JSON gives the assistant named fields it can inspect, compare, and pass to the next step.\n• The harness can validate the same fields instead of trusting a confident summary.\n• Errors should use the same predictable shape as successes.\n• Transition: now watch those structured results drive a complete state change.",["C:/Users/woody/Workspace/vault/wiki/pncli/pncli-overview.md"]);
}

// 6 — demo setup
{
  const s=pres.slides.add(); chrome(s,"05 · live demo",6,12,"THE BUG THAT FILES ITSELF");
  kicker(s,"Watch the state change");
  txt(s,"A ticket becomes evidence,\nthen action.",{left:64,top:145,width:980,height:128},{size:56,bold:true,lineSpacing:.93});
  const beats=[
    ["01","FIND","Read the ticket","URL · steps · owner",C.violet],
    ["02","PROVE","Drive the page","Failure · console · screenshot",C.cyan],
    ["03","GATE","Ask once","Approve the exact write-back",C.amber],
    ["04","ROUTE","Update Jira","Attach · comment · assign · move",C.cyan]
  ];
  beats.forEach((b,i)=>{const x=64+i*288; rule(s,x,355,245,b[4],4); txt(s,b[0],{left:x,top:382,width:40,height:18},{size:11,bold:true,color:b[4],mono:true}); txt(s,b[1],{left:x,top:425,width:240,height:34},{size:22,bold:true,color:b[4]}); txt(s,b[2],{left:x,top:475,width:250,height:34},{size:25,bold:true}); txt(s,b[3],{left:x,top:532,width:244,height:44},{size:16,color:C.muted});});
  note(s,"Talking points:\n• Set up the scenario: a Jira ticket contains a URL and reproduction steps.\n• FIND is read-only: retrieve the ticket and identify the target state.\n• PROVE uses the browser: reproduce the failure, capture the console error, and take a screenshot.\n• GATE is the hinge: stop visibly before the first write.\n• ROUTE writes the evidence back, assigns an owner, and changes the ticket state.\n• Keep a prepared issue and deterministic bad input so the demonstration is repeatable.\n• Transition: zoom in on the moment where observation becomes authority.",["C:/Users/woody/Workspace/vault/documents/presentations/demo/demo-ideas.md"]);
}

// 7 — permission hinge
{
  const s=pres.slides.add(); chrome(s,"06 · the permission moment",7,12,"HUMAN AUTHORITY STAYS VISIBLE");
  kicker(s,"The demo's hinge",64,105,C.amber);
  txt(s,"Reading was automatic.\nWriting is a decision.",{left:64,top:165,width:570,height:150},{size:54,bold:true,lineSpacing:.92});
  txt(s,"Approve a bounded action—not a vague promise of good behavior.",{left:68,top:365,width:500,height:82},{size:25,color:"#D1D4DA"});
  shp(s,"roundRect",{left:700,top:120,width:516,height:485},"#18140E",C.amber,1);
  txt(s,"COPILOT WANTS TO RUN",{left:736,top:156,width:350,height:20},{size:12,bold:true,color:C.muted,mono:true});
  txt(s,"pncli jira add-comment\n--key APP-901 …",{left:736,top:215,width:430,height:78},{size:24,bold:true,mono:true});
  const opts=[["1 · Run this command",true],["2 · Approve this tool for the session",false],["3 · No, change the approach",false]];
  opts.forEach((o,i)=>{shp(s,"roundRect",{left:736,top:345+i*72,width:440,height:54},o[1]?"#2A2115":C.panel,o[1]?C.amber:C.line,1);txt(s,o[0],{left:756,top:360+i*72,width:400,height:24},{size:17,bold:o[1],color:o[1]?C.amber:C.muted,mono:true});});
  note(s,"Talking points:\n• Reading happened without ceremony because it did not change the system of record.\n• The proposed write is concrete: one command, one ticket, one visible effect.\n• Read the prompt aloud before approving it; demonstrate what informed consent looks like.\n• Do not choose session-wide approval—the lesson is bounded authority, not convenience.\n• A permission prompt is a product surface, not an interruption to hide.\n• Transition: approval is only meaningful if the action leaves a verifiable receipt.",["C:/Users/woody/Workspace/vault/wiki/copilot/copilot-cli-permissions-and-sandboxing.md"]);
}

// 8 — receipt
{
  const s=pres.slides.add(); chrome(s,"07 · the receipt",8,12,"PROOF, NOT VIBES");
  kicker(s,"How we know it worked");
  txt(s,"The result leaves a trail another human can inspect.",{left:64,top:145,width:1080,height:102},{size:49,bold:true,lineSpacing:.95});
  const rows=[
    ["REPORTED","To Do · needs-repro",C.muted],
    ["OBSERVED","Checkout fails · console error captured",C.red],
    ["EVIDENCE","Screenshot + exact reproduction note",C.cyan],
    ["AUTHORIZED","Comment · assignment · transition",C.amber],
    ["FINAL","In Progress · owner set",C.cyan]
  ];
  rows.forEach((r,i)=>{const y=310+i*60; rule(s,64,y,1152,C.line); txt(s,r[0],{left:64,top:y+18,width:210,height:20},{size:13,bold:true,color:C.muted,mono:true}); txt(s,r[1],{left:350,top:y+12,width:866,height:28},{size:20,bold:true,color:r[2],mono:true,align:"right"});});
  note(s,"Talking points:\n• Reveal the updated ticket only after the run completes.\n• Trace each row: reported state, observed state, captured evidence, authorized mutation, final state.\n• The screenshot alone is not success; the system of record must reflect what was observed.\n• Another human should be able to reconstruct what happened without trusting the assistant’s narration.\n• This is evaluation in practical form: compare intended, observed, and final states.\n• Transition: sometimes the correct evaluated result is to refuse the requested action.");
}

// 9 — refusal
{
  const s=pres.slides.add(); chrome(s,"08 · evaluation",9,12,"REFUSAL CAN BE THE RIGHT RESULT");
  kicker(s,"The stronger demo");
  rich(s,[{run:"Give it the merge button.\nThen watch it ",bold:true},{run:"refuse",bold:true,color:C.red},{run:".",bold:true}],{left:64,top:170,width:1120,height:190},{size:62,lineSpacing:.94});
  shp(s,"rect",{left:64,top:445,width:10,height:116},C.red,"none",0);
  txt(s,"Sonar fails. CI is red. Security has an open finding.\nA useful agent preserves the conditions under which completion is allowed.",{left:100,top:438,width:1020,height:120},{size:28,color:"#D2D4D9",lineSpacing:1.15});
  note(s,"Talking points:\n• Giving an agent a merge tool does not mean every task should end with a merge.\n• Name the independent checks: Sonar, CI, and security findings.\n• The agent’s job is to preserve the conditions for safe completion, not merely maximize task completion.\n• A refusal with evidence is a successful outcome.\n• If time permits, use demo idea 3 and let the audience predict the decision before revealing it.\n• Transition: powerful tools make the origin and behavior of skills a supply-chain concern.",["C:/Users/woody/Workspace/vault/documents/presentations/demo/demo-ideas.md"]);
}

// 10 — provenance visual
{
  const s=pres.slides.add(); s.background.fill=C.bg;
  await addImage(s,`${ASSET}/provenance-gate.png`,{left:0,top:0,width:1280,height:720},"A software artifact passing through transparent provenance and verification layers");
  shp(s,"rect",{left:0,top:0,width:535,height:720},"#07090D","none",0);
  txt(s,"09 · SKILLS MARKETPLACES",{left:64,top:42,width:430,height:20},{size:12,bold:true,color:C.cyan,mono:true});
  kicker(s,"Power needs provenance",64,145);
  txt(s,"A skill that steers tools\nis part of your supply chain.",{left:64,top:200,width:510,height:158},{size:49,bold:true,lineSpacing:.94});
  txt(s,"Pin the source. Read the ask. Constrain the run.",{left:68,top:430,width:430,height:86},{size:24,color:"#D1D4DA",lineSpacing:1.15});
  note(s,"Talking points:\n• A skill can shape how an assistant uses tools, so installation is part of the execution supply chain.\n• Pin the source: know the repository, publisher, and exact revision.\n• Read the ask: compare requested files, network, and commands with the job the skill claims to perform.\n• Constrain the run: use narrow permissions, sandboxing, and an audit trail.\n• Scanning helps, but provenance and least privilege remain essential.\n• Transition: compress the session into three operating habits.",["C:/Users/woody/Workspace/vault/wiki/agent-harness-engineering/agent-skills-marketplaces.md","C:/Users/woody/Workspace/vault/wiki/agent-harness-engineering/agent-skills-supply-chain-security.md","AI-generated visual: OpenAI image generation, created for this deck"]);
}

// 11 — operating principles
{
  const s=pres.slides.add(); chrome(s,"10 · operating principles",11,12,"MAKE AUTHORITY LEGIBLE");
  kicker(s,"Three habits to take back to work");
  txt(s,"Let reads flow.\nMake writes specific.\nDemand a receipt.",{left:64,top:150,width:860,height:260},{size:62,bold:true,lineSpacing:1.02});
  const labels=[["READ",C.violet],["AUTHORIZE",C.amber],["PROVE",C.cyan]];
  labels.forEach((l,i)=>{const x=64+i*382;rule(s,x,520,330,l[1],4);txt(s,l[0],{left:x,top:548,width:330,height:26},{size:18,bold:true,color:l[1],mono:true});});
  note(s,"Talking points:\n• Let reads flow when the risk is understood and the scope is appropriate.\n• Make writes specific: identify the exact operation, target, and expected effect.\n• Demand a receipt: capture the result in a form another person can inspect.\n• These habits are vendor-independent and apply to terminal agents, workplace assistants, and custom automations.\n• Invite the audience to test one workflow with these three questions this week.\n• Transition: close by turning the three habits into three memorable questions.");
}

// 12 — closing
{
  const s=pres.slides.add(); chrome(s,"11 · take it back to work",12,12,"THE QUESTION TO KEEP");
  kicker(s,"Close where we opened");
  rich(s,[{run:"What did it ",bold:true},{run:"read",bold:true,color:C.cyan},{run:"?",bold:true}],{left:64,top:165,width:1100,height:74},{size:59});
  rich(s,[{run:"What did it ",bold:true},{run:"do",bold:true,color:C.violet},{run:"?",bold:true}],{left:64,top:265,width:1100,height:74},{size:59});
  rich(s,[{run:"Who said ",bold:true},{run:"yes",bold:true,color:C.amber},{run:"?",bold:true}],{left:64,top:365,width:1100,height:74},{size:59});
  rule(s,64,505,86,C.cyan,5);
  txt(s,"Good tools make action possible.\nGood harnesses make authority visible.",{left:64,top:545,width:950,height:70},{size:28,color:"#D3D5D9",lineSpacing:1.12});
  note(s,"Talking points:\n• Ask the three questions slowly and leave space after each one.\n• ‘What did it read?’ identifies context, permissions, and possible exposure.\n• ‘What did it do?’ separates a generated answer from a real state change.\n• ‘Who said yes?’ reveals the policy or person that supplied authority.\n• Invite one example from the room and classify it aloud: read, propose, authorize, act, prove.\n• Close on the final line: good tools make action possible; good harnesses make authority visible.");
}

await fs.mkdir(QA,{recursive:true});
await fs.mkdir(path.dirname(OUT),{recursive:true});
for(const [i,s] of pres.slides.items.entries()){
  const stem=`slide-${String(i+1).padStart(2,"0")}`;
  const png=await pres.export({slide:s,format:"png",scale:1});
  await fs.writeFile(path.join(QA,stem+".png"),new Uint8Array(await png.arrayBuffer()));
  const layout=await s.export({format:"layout"});
  await fs.writeFile(path.join(QA,stem+".layout.json"),await layout.text());
}
const montage=await pres.export({format:"webp",montage:true,scale:1});
await fs.writeFile(path.join(QA,"montage.webp"),new Uint8Array(await montage.arrayBuffer()));
const pptx=await PresentationFile.exportPptx(pres);
await pptx.save(OUT);
