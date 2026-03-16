# 📖 FileNet & ICN Stored Search Architecture

This repository contains documentation and PowerShell tooling for dynamically generating and deploying complex stored searches in IBM FileNet P8, IBM Content Navigator (ICN), and Workplace XT.

## 🔗 Important Reference
* **IBM Schema Reference:** [FileNet P8 Platform 5.6.0 Schema Reference](https://www.ibm.com/docs/en/filenet-p8-platform/5.6.0?topic=roadmap-schema-reference)

## 🏗️ System Limitations & Quirks (ICN & Workplace XT)
Both the modern interface (ICN) and the legacy interface (Workplace XT) have hardcoded UI rendering limitations when it comes to complex mathematical queries, particularly nested `OR` conditions. To bypass this, we use a "Split-Brain" approach: a simple query for the UI, and a complex mathematical query for the backend engine.

### File Structures & Storage
When a search is saved in FileNet, it is stored in a `StoredSearch` object. The actual search logic and layout are saved as *Content Elements* (attachments) on this object.
* **ICN (IBM Content Navigator)** creates two files:
  * `file0` (`application/x-filenet-searchtemplate`): The XML backend logic.
  * `file1` (`application/json`): The UI layout and behavior.
* **Workplace XT** creates two XML files:
  * `Content.xml`: The backend logic.
  * `UIContent.xml`: The UI layout.

*Note: You can deploy an ICN search with only `file0` (XML) to the `StoredSearch` object. ICN will attempt to render a UI based on the XML alone, but this must be thoroughly tested as ICN often flattens complex binary logic.*

### Search Concepts
* **Search Templates:** Think of these as blank forms where the user fills in the criteria.
* **Stored Searches:** Pre-defined queries where criteria are fixed (or partially editable). Logic is stored in `<whereprop>` tags.
* **SearchConfiguration:** An ICN-specific object for desktop layouts and settings. It has *nothing* to do with mathematical query logic. (Note: Extensions like `.itms` or `.vhds` belong to other archiving solutions and are not part of the FileNet search architecture).

## 🧠 Logic & UI Rendering Rules
* **ICN Base Group (Global Group):** Attributes within the base group share a single fixed `AND` or `OR` operator.
* **ICN Multiple Groups:** The relationship between different groups (e.g., Group 1 `AND` Group 2) strictly inherits the operator chosen in the base group. ICN *cannot* render mixed logic (e.g., an `OR` between groups if the base group uses `AND`).
* **XT Compatibility:** ICN can open a Workplace XT search. It maps parameters to the base group, but handles all other XT groups as pre-filled, hidden background logic.

## ⚙️ Backend (CPE) API & XML Rules
The Content Platform Engine (CPE) requires the XML to be built very specifically:
* **Binary XML (The "Russian Doll"):** The `<where>` clause does *not* accept flat lists for `AND`/`OR`. It strictly compares a maximum of 2 attributes or groups at a time. This creates a deep, binary tree:
```text
Root <and>
 ├── Global Attribute 1 (Editable)
 └── <and>
      ├── Global Attribute 2 (Editable)
      └── <and>
           ├── Global Attribute 3 (Editable)
           └── <or>  <-- Start of the OR block
                ├── Group 1 (A+B+C)
                └── <or>
                     ├── Group 2 (A+B+C)
                     └── <or>
                          ├── Group 3 (A+B+C)
                          └── <or>
                               ├── Group 4 (A+B+C)
                               └── Group 5 (A+B+C)  <-- Final binary pair
```
* **No Direct SQL:** FileNet SQL cannot be injected directly into the XML. The engine strictly parses the `<whereprop>` tags.
* **Operator Translation:** JSON and XML use different operators. For example, JSON uses `EQUAL` or `GREATEROREQUAL`, while XML uses `eq` or `gte`.
* **Specials:** Be aware of operators like `INANY` (XML: `in`) and `NOTIN` (XML: `neq`).
* **CDATA Wrapping:** Values in the XML must be wrapped in `<![CDATA[...]]>` to prevent special characters (`&`, `<`, `>`) from breaking the parser.
* **Avoid Smartoperators:** Do not use `<smartoperator>` tags; they are poorly documented and cause instability.
* **Internal Engine:** ICN's internal engine automatically generates the `<executedata>` XML block for you during execution; you do not need to script this manually.

## 🚀 The PowerShell Automation Tool
Manually editing these structures via ACCE is highly error-prone. To use the included PowerShell tool:
1. Create a (dummy) search in the ICN UI and save it (so the `StoredSearch` object and GUID exist in FileNet).
2. Use the PS1 script to dynamically calculate the Cartesian Blacklist, generate the required XML/JSON/SQL, and automatically replace the content elements via the CPE API.
