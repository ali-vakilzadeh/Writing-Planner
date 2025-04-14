# Writing Planner Basic - Word Add-in

Welcome to the Writing Planner Basic! This Word Add-in helps you to structure, plan, and track your writing projects directly within Microsoft Word. Say goodbye to scattered notes and hello to organized, efficient writing!

## Project Description

Writing Planner Basic is a task pane add-in for Microsoft Word, built with React, Fluent UI, and Tailwind CSS. It provides a dedicated space to outline your document, manage the status of each section, add notes, and visualize your overall progress.

**Key Features:**

*   **Document Planning:** Create a hierarchical structure for your document sections (chapters, topics, etc.).
*   **Status Tracking:** Assign a status to each section (e.g., Empty, Drafted, Edited, Finalized) to monitor progress.
*   **Progress Visualization:** See the overall completion percentage based on section statuses.
*   **Section Comments:** Add notes, reminders, or specific instructions to each section.
*   **Document Synchronization:**
    *   **Build Document:** Generate basic heading structures in your Word document based on your plan.
    *   **Sync Plan:** Update your plan by detecting existing headings within your document.
*   **Statistics:** Get insights into section content (word count, paragraphs, etc. - requires document interaction).
*   **Default Template:** Start quickly with a pre-defined template for common document structures (loads automatically in empty documents).
*   **Data Persistence:** Your plan is saved directly within the Word document's properties.
*   **Import/Export TOC:** Share or reuse your document structures by importing/exporting the Table of Contents definition.

## Installation Guide

Get ready to supercharge your writing workflow! Follow these steps to install the Writing Planner Basic add-in:

**Prerequisites:**

*   **Node.js and npm/pnpm:** Ensure you have Node.js installed. This project uses pnpm (`pnpm-lock.yaml` found), but npm should also work. You can get Node.js (which includes npm) from [nodejs.org](https://nodejs.org/). Install pnpm globally if needed: `npm install -g pnpm`.
*   **Microsoft Word:** You need a version of Microsoft Word that supports Office Add-ins (Microsoft 365 is recommended).
*   **Git:** To clone the repository.

**Steps:**

1.  **Clone the Repository:**
    ```bash
    git clone https://github.com/OfficeDev/Office-Addin-TaskPane-React-JS.git # Or your specific repo URL
    cd office-addin-taskpane-react-js # Or your project directory
    ```

2.  **Install Dependencies:**
    Open your terminal in the project directory and run:
    ```bash
    pnpm install
    # OR if using npm:
    # npm install
    ```

3.  **Trust Development Certificates:**
    Office Add-ins require HTTPS. You might need to trust the development certificate:
    ```bash
    pnpm run install-dev-certs # Check package.json for the exact script if different
    # OR if using office-addin-cli directly:
    # npx office-addin-dev-certs install
    ```
    Follow the prompts to trust the certificate.

4.  **Build the Add-in:**
    Compile the add-in code:
    ```bash
    pnpm run build
    # OR
    # npm run build
    ```

5.  **Sideload the Add-in into Word:**
    This command starts the development server and automatically sideloads the add-in into Word (or your specified Office application in `package.json`).
    ```bash
    pnpm run start
    # OR
    # npm run start
    ```
    This will likely open Word for you with the add-in loaded. If Word doesn't open automatically, open it manually.

6.  **Find the Add-in:** The add-in task pane should appear automatically. If not, you can usually find sideloaded add-ins under the "Home" or "Insert" tab in Word (look for a button with the add-in's name or a generic "Office Add-ins" button).

## User Guide

Let's get planning!

**Opening the Add-in:**

*   Once installed and sideloaded (using `pnpm start`), the "Writing Planner Basic" task pane should appear on the right side of your Word window.
*   If you close it, you can typically reopen it from the "Home" tab in the Word ribbon.

**Using the Planner:**

The task pane has two main tabs: "Plan" and "TOC".

**1. Plan Tab:**

*   **Viewing the Plan:** This tab shows your document structure as a list of sections.
*   **Adding Sections:** Click the **"Add Section"** button to add a new top-level section. You might be able to add sub-sections depending on the implementation.
*   **Editing Titles:** Click on a section's title to edit it.
*   **Changing Status:** Use the dropdown menu next to each section to update its status (Empty, Drafted, Finalized, etc.). The color coding provides a quick visual cue.
*   **Adding Comments:** Click the comment icon (💬) next to a section to open a panel where you can add or edit notes for that section.
*   **Viewing Statistics:** Click the statistics icon (📊) to see details like word count for the content associated with that section in the document (requires refresh).
*   **Deleting Sections:** Click the delete icon (🗑️) to remove a section from your plan.
*   **Overall Progress:** The progress bar at the top shows the calculated completion percentage based on the status of all sections.

**2. TOC (Table of Contents) Tab:**

*   **Viewing the TOC:** This tab displays the planned structure, often resembling a Table of Contents.
*   **Import/Export:** Use the **"Import TOC"** and **"Export TOC"** buttons to load or save your document structure plan as a file. This is great for templates or sharing outlines.

**Core Actions (Buttons at the top/bottom):**

*   **Refresh Stats:** Click this to re-scan the document and update the word counts and other statistics for each section based on the current content linked to the headings.
*   **Sync Plan:** Click this to analyze the headings (using Word's built-in Heading styles) in your document and add any missing ones to your plan.
*   **Build Document:** Click this to insert basic headings into your Word document based on the sections currently defined in your plan. *Use with caution, especially in existing documents.*
*   **Delete All Data:** Click this (often requires confirmation) to completely remove the planner data from the current document's properties.

**Saving:**

Your plan is automatically saved within the document's properties whenever you make changes. Just save your Word document as usual!

---

Happy Writing! We hope the Writing Planner Basic helps streamline your creative process.