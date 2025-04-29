# AI Development Guide - FDA 510(k) Intelligence Suite

## 1. Purpose

This document outlines the process, context, and standard prompts used when collaborating with AI assistants (like Google Gemini, GitHub Copilot Chat, ChatGPT) during the development, debugging, and refinement of the Excel VBA-based FDA 510(k) Intelligence Suite. Its goal is to ensure consistency, provide necessary context for effective AI interaction, and streamline future AI-assisted development efforts.

## 2. Core Architecture Summary for AI Context

When prompting an AI assistant, provide this summary:

"The project is an FDA 510(k) Lead Scoring tool built entirely within a single Excel `.xlsm` workbook. It uses:
* **Power Query (M):** Fetches previous month's 510(k) data from the openFDA API dynamically. Code likely in `src/powerquery/FDA_510k_Query.pq`.
* **VBA Modules:**
    * `ThisWorkbook`: Handles `Workbook_Open` trigger, initiates refresh and processing.
    * `mod_510k_Processor`: Main orchestration module. Contains `ProcessMonthly510k` (workflow control), `Calculate510kScore` (core scoring logic), caching functions (`GetCompanyRecap`, `Load/SaveCompanyCache`), parameter loading (`LoadWeightsAndKeywords`), formatting routines (`ReorganizeColumns`, `FormatTableLook`, etc.), archiving (`ArchiveMonth`), and various helpers. Interacts heavily with worksheet tables and data arrays.
    * `mod_Logger`: Provides a performance-optimized, buffered logging system (`LogEvt`, `FlushLogBuf`) writing to a hidden `RunLog` sheet.
* **Excel Sheets:** `CurrentMonthData` (PQ output, final results), `Weights` (parameter tables like `tblACWeights`, `tblKeywords`, `tblNFCosmeticKeywords`, etc.), `CompanyCache` (persistent recap storage), `RunLog` (hidden log).
* **Key Libraries:** VBA references include `Scripting Runtime` (Dictionary, FileSystemObject), `Microsoft XML, v6.0` (ServerXMLHTTP for OpenAI), `WScript.Shell` (Environment Strings).
* **Goal:** Automate fetching, scoring (based on configurable weights/rules/keywords), caching, formatting, and archiving 510(k) leads, providing results directly in Excel."

## 3. Key Development Stages & AI Interaction Examples

* **Initial Design:** Provided high-level requirements, constraints (initially Excel-only, later VBA allowed), desired workflow (PQ -> VBA Score -> Cache -> Archive), and data inputs/outputs. AI helped generate initial VBA function/sub skeletons and module structure.
* **Power Query:** Provided draft M code or described requirements (dynamic dates, specific API endpoint, fields needed, calculations like ProcTimeDays). Asked AI to refine M code for robustness, error handling, and efficiency.
* **VBA - Scoring Logic:** Provided scoring model details (weights AC, PC, KW, etc.; NF rules based on keywords; Synergy rules). Asked AI to implement `Calculate510kScore` function using lookups (dictionaries loaded from tables), keyword checks (collections), and the specified formula structure. *Crucial: AI needs the exact formula and NF/Synergy rules confirmed by the developer.*
* **VBA - Caching & API:** Described caching requirement (in-memory dictionary, persistent sheet). Described optional OpenAI recap requirement (maintainer only, secure API key handling via `%APPDATA%`). Asked AI to implement `Load/SaveCompanyCache`, `GetCompanyRecap`, `GetCompanyRecapOpenAI` (including HTTP request structure). *Crucial: AI needs guidance on prompt engineering for OpenAI and acknowledgement that basic parsing needs replacement.*
* **VBA - Workflow & Formatting:** Described desired `Workbook_Open` behavior, Day Guard logic, Archive logic, and specific final sheet formatting (column order, styles, borders, conditional colors, comments for long text, freeze panes). Asked AI to implement `ProcessMonthly510k` orchestration, `ReorganizeColumns`, `FormatTableLook`, `FormatCategoryColors`, `CreateShortNamesAndComments`, `FreezeHeaderAndKeyCols` and ensure correct call order.
* **Debugging:** Provided specific compile error messages and screenshots (`AddComment` `Threaded` arg, duplicate declaration, `Next without For`, `End If without Block If`). Asked AI to identify the cause and provide corrected code snippets.
* **Logging:** Requested robust logging solutions. Evaluated AI suggestions (simple vs. buffered). Asked AI to generate the `mod_Logger` code and integrate `LogEvt`/`FlushLogBuf` calls throughout the project.
* **Documentation:** Asked AI to generate `.gitignore`, `README.md`, `ARCHITECTURE.md`, and this guide based on the project context and code.
* **Version Control:** Used AI to generate `git commit` messages summarizing changes.

## 4. Standard Prompts & Context Provision

Effective AI collaboration requires clear context:

* **Role & Goal:** Start prompts with "Act as an expert VBA developer..." and state the specific task (e.g., "review this function for errors", "implement this logic", "refactor this code for performance").
* **Provide Context:** Use the Architecture Summary (Section 2) or relevant parts. Explain *what* the specific function/module does in the overall workflow.
* **Provide Code:** Paste the relevant VBA function, module (`.bas`/`.cls`), or M query (`.pq`) directly into the prompt or ensure the file is open in the VS Code editor if using integrated AI.
* **Be Specific:**
    * For errors: Provide the *exact* error message and the highlighted code line/block.
    * For new features: Clearly define inputs, outputs, logic steps, and any constraints. Reference specific functions or variables to modify.
    * For reviews: State *what* to review for (e.g., correctness, performance, error handling, best practices, adherence to requirements).
* **Iterate:** Treat it as a conversation. Provide feedback on the AI's suggestions, ask clarifying questions, provide corrected information if the AI misunderstands.

**Example Review Prompt:** (See previous detailed prompt generated for VS Code AI)

## 5. Handling AI Limitations

* **VBA Nuances:** AI might not always grasp subtle VBA object model behaviors, error handling nuances (`On Error`), or Excel-specific interactions. Developer oversight is crucial.
* **Context Window:** Large codebases might exceed the AI's context window. Focus prompts on specific modules or functions. Provide summaries of related code if necessary.
* **"Hallucinations":** AI might invent functions, methods, or arguments (like the `Threaded` argument for `AddComment` in older Excel). Always test AI-generated code thoroughly.
* **Specificity:** AI performs best with specific instructions. Vague requests ("make it better") yield generic results.

## 6. VBA Export/Import for Git & AI

* To use Git effectively and provide clean code context to AI assistants, VBA code should be exported from the Excel VBE:
    * Right-click module/class/form/sheet in VBE Project Explorer -> `Export File...`
    * Save as `.bas` (modules), `.cls` (classes, ThisWorkbook, Sheets), `.frm` (forms).
    * Commit these text files to Git.
* Provide the content of these exported files to the AI.
* To import updated code:
    * Remove the existing module/class/etc. from the VBE project.
    * `File` -> `Import File...` -> Select the updated `.bas`/`.cls`/`.frm` file.

## 7. Future Development

* Use this guide and the `ARCHITECTURE.md` to brief AI assistants on the existing system before requesting modifications or new features.
* Provide relevant sections of existing code when asking for changes.
* Document new features or major changes in the `README.md` and `ARCHITECTURE.md`, potentially updating this guide as well.
* Refer to the `RunLog` sheet for diagnosing issues before consulting AI.