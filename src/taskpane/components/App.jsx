"use client"
import * as React from "react"
import { useState, useEffect, useRef } from "react"
import { DefaultButton, PrimaryButton, IconButton } from "@fluentui/react/lib/Button"
import { Stack } from "@fluentui/react/lib/Stack"
import { TextField } from "@fluentui/react/lib/TextField"
import { Dropdown } from "@fluentui/react/lib/Dropdown"
import { Pivot, PivotItem } from "@fluentui/react/lib/Pivot"
import { ProgressIndicator } from "@fluentui/react/lib/ProgressIndicator"
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog"
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner"
import { TooltipHost } from "@fluentui/react/lib/Tooltip"
import { initializeIcons } from "@fluentui/react/lib/Icons"
import { mergeStyles } from "@fluentui/react/lib/Styling"
import { Callout } from "@fluentui/react/lib/Callout"
import { Label } from "@fluentui/react/lib/Label"
import { Panel } from "@fluentui/react/lib/Panel"
import { Word } from "../utils/mock-office"

// Initialize icons
initializeIcons()

// Status options for document sections
const STATUS_OPTIONS = [
  { key: "empty", text: "Empty" },
  { key: "created", text: "Created" },
  { key: "drafted", text: "Drafted" },
  { key: "checked", text: "Checked" },
  { key: "referenced", text: "Referenced" },
  { key: "edited", text: "Edited" },
  { key: "verified", text: "Verified" },
  { key: "finalized", text: "Finalized" },
]

// Assign partial progress to incomplete statuses
const STATUS_PROGRESS = {
  empty: 0,
  created: 5,
  drafted: 20,
  checked: 30,
  referenced: 50,
  edited: 70,
  verified: 90,
  finalized: 100,
}

// Status colors for visual indication
const STATUS_COLORS = {
  empty: { background: "#f3f2f1", color: "#605e5c" },
  created: { background: "#deecf9", color: "#2b88d8" },
  drafted: { background: "#fff4ce", color: "#c19c00" },
  checked: { background: "#dff6dd", color: "#107c10" },
  referenced: { background: "#f3e9ff", color: "#8764b8" },
  edited: { background: "#e5e5ff", color: "#5c5cc0" },
  verified: { background: "#e0f5f5", color: "#038387" },
  finalized: { background: "#d0f0e0", color: "#0b6a0b" },
}

// Styles
const containerStyles = {
  root: {
    padding: 10,
    width: "100%",
    height: "100%",
    boxSizing: "border-box",
    overflow: "auto",
  },
}

const headerStyles = {
  root: {
    padding: "10px 0",
    borderBottom: "1px solid #edebe9",
  },
}

const titleStyles = {
  root: {
    fontSize: 18,
    fontWeight: 600,
    margin: 0,
  },
}

const subtitleStyles = {
  root: {
    fontSize: 12,
    color: "#605e5c",
    margin: "4px 0 0 0",
  },
}

const sectionNameStyle = mergeStyles({
  cursor: "pointer",
  fontWeight: 500,
  fontSize: "13px",
  ":hover": {
    textDecoration: "underline",
  },
})

const sectionRowStyle = mergeStyles({
  display: "flex",
  justifyContent: "space-between",
  alignItems: "center",
  width: "100%",
  padding: "4px 0",
})

const buttonRowStyle = mergeStyles({
  display: "flex",
  justifyContent: "flex-start",
  alignItems: "center",
  width: "100%",
  padding: "2px 0",
  marginLeft: "8px",
})

const statusCellClass = mergeStyles({
  textAlign: "center",
  padding: "2px 4px",
  borderRadius: 2,
  fontSize: 12,
  fontWeight: 600,
  display: "inline-block",
  minWidth: 70,
})

export default function App(props) {
  const { isOfficeInitialized = true } = props || {}
  const containerRef = useRef(null)
  const [containerWidth, setContainerWidth] = useState(0)

  const [tocItems, setTocItems] = useState([])
  const [planningItems, setPlanningItems] = useState([])
  const [activeTab, setActiveTab] = useState("plan")
  const [refreshing, setRefreshing] = useState(false)
  const [nextId, setNextId] = useState(1)
  const [dataLoaded, setDataLoaded] = useState(false)
  const [templateApplied, setTemplateApplied] = useState(false)
  const [aboutOpen, setAboutOpen] = useState(false)
  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false)
  const [buildingToc, setBuildingToc] = useState(false)
  const [buildingDocument, setBuildingDocument] = useState(false)
  const [editingItem, setEditingItem] = useState(null)
  const [commentItem, setCommentItem] = useState(null)
  const [statsCalloutVisible, setStatsCalloutVisible] = useState(false)
  const [statsCalloutTarget, setStatsCalloutTarget] = useState(null)
  const [statsItem, setStatsItem] = useState(null)
  const [error, setError] = useState(null)
  const [documentIsEmpty, setDocumentIsEmpty] = useState(true)

  // Monitor window resize for responsive layout
  useEffect(() => {
    const handleResize = () => {
      if (containerRef.current) {
        setContainerWidth(containerRef.current.clientWidth)
      }
    }

    // Set initial width
    handleResize()

    // Add event listener
    window.addEventListener("resize", handleResize)

    // Clean up
    return () => {
      window.removeEventListener("resize", handleResize)
    }
  }, [])

  // Load data from document properties
  useEffect(() => {
    if (isOfficeInitialized && !dataLoaded) {
      loadFromDocumentProperties()
    }
  }, [isOfficeInitialized, dataLoaded])

  // Initialize with template structure if nothing was loaded
  useEffect(() => {
    if (dataLoaded && tocItems.length === 0 && planningItems.length === 0 && !templateApplied) {
      createTemplateStructure()
      setTemplateApplied(true)
    }
  }, [dataLoaded, tocItems.length, planningItems.length, templateApplied])

  // Check if document is empty
  const checkIfDocumentIsEmpty = async () => {
    try {
      if (!Word || typeof Word.run !== "function") {
        console.log("Word API not available, assuming empty document for development")
        setDocumentIsEmpty(true)
        return true
      }

      let isEmpty = false
      await Word.run(async (context) => {
        const body = context.document.body
        body.load("text")
        await context.sync()

        // If document has no text or only whitespace, consider it empty
        isEmpty = !body.text || body.text.trim().length === 0
      })

      setDocumentIsEmpty(isEmpty)
      return isEmpty
    } catch (error) {
      console.error("Error checking if document is empty:", error)
      return true // Assume empty on error
    }
  }

  const loadFromDocumentProperties = async () => {
    try {
      // First check if document is empty
      const isEmpty = await checkIfDocumentIsEmpty()

      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")

        // Try to get from localStorage for development
        const plannerData = localStorage.getItem("documentPlannerData")

        if (plannerData) {
          try {
            const data = JSON.parse(plannerData)

            if (data && data.tocItems && data.planningItems) {
              setTocItems(data.tocItems || [])
              setPlanningItems(
                (data.planningItems || []).map((item) => ({
                  ...item,
                  words: 0,
                  paragraphs: 0,
                  tables: 0,
                  graphics: 0,
                })),
              )

              // Find the highest ID to set nextId correctly
              const highestId = Math.max(...data.planningItems.map((item) => item.id || 0), 0)
              setNextId(highestId + 1)

              // Refresh statistics after loading data
              setTimeout(() => refreshStatistics(), 500)
            }
          } catch (parseError) {
            console.error("Error parsing planner data:", parseError)
          }
        }

        setDataLoaded(true)
        return
      }

      await Word.run(async (context) => {
        try {
          // Get document properties
          const properties = context.document.properties.customProperties
          if (!properties) {
            console.error("Document properties not available")
            setDataLoaded(true)
            return
          }

          properties.load("key,value")
          await context.sync()

          // Find our data property
          let plannerData = null
          if (properties.items && Array.isArray(properties.items)) {
            for (let i = 0; i < properties.items.length; i++) {
              if (properties.items[i] && properties.items[i].key === "documentPlannerData") {
                plannerData = properties.items[i].value
                break
              }
            }
          } else {
            // Try to get from localStorage for development
            plannerData = localStorage.getItem("documentPlannerData")
          }

          if (plannerData) {
            try {
              const data = JSON.parse(plannerData)

              if (data && data.tocItems && data.planningItems) {
                // Find the highest ID to set nextId correctly
                const highestId = Math.max(...data.planningItems.map((item) => item.id || 0), 0)

                setTocItems(data.tocItems || [])

                // Set planning items but initialize statistics to 0
                setPlanningItems(
                  (data.planningItems || []).map((item) => ({
                    ...item,
                    words: 0,
                    paragraphs: 0,
                    tables: 0,
                    graphics: 0,
                  })),
                )

                setNextId(highestId + 1)

                // Refresh statistics after loading data
                setTimeout(() => refreshStatistics(), 500)
              }
            } catch (parseError) {
              console.error("Error parsing planner data:", parseError)
            }
          } else if (isEmpty) {
            // If document is empty and no saved data, we'll create a template later
            console.log("Document is empty and no saved data found")
          }

          setDataLoaded(true)
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)
          setDataLoaded(true)
        }
      })
    } catch (error) {
      console.error("Error loading data:", error)
      // Fallback to empty arrays if there's an error
      //setTocItems([])
      //setPlanningItems([])
      setDataLoaded(true)
      setError("Failed to load data. Please try again.")
    }
  }

  const createTemplateStructure = () => {
    try {
      // Comprehensive document template structure
      const templateTocItems = [
        { id: 1, title: "Title Page", level: 1, isDefault: true },
        { id: 2, title: "Abstract", level: 1, isDefault: true },
        { id: 3, title: "Table of Contents", level: 1, isDefault: true },
        { id: 4, title: "List of Figures", level: 1, isDefault: true },
        { id: 5, title: "List of Tables", level: 1, isDefault: true },
        { id: 6, title: "Introduction", level: 1, isDefault: true },
        { id: 7, title: "Background", level: 2, isDefault: true },
        { id: 8, title: "Problem Statement", level: 2, isDefault: true },
        { id: 9, title: "Research Questions", level: 2, isDefault: true },
        { id: 10, title: "Significance of Study", level: 2, isDefault: true },
        { id: 11, title: "Literature Review", level: 1, isDefault: true },
        { id: 12, title: "Theoretical Framework", level: 2, isDefault: true },
        { id: 13, title: "Previous Research", level: 2, isDefault: true },
        { id: 14, title: "Research Gap", level: 2, isDefault: true },
        { id: 15, title: "Methodology", level: 1, isDefault: true },
        { id: 16, title: "Research Design", level: 2, isDefault: true },
        { id: 17, title: "Data Collection", level: 2, isDefault: true },
        { id: 18, title: "Data Analysis", level: 2, isDefault: true },
        { id: 19, title: "Ethical Considerations", level: 2, isDefault: true },
        { id: 20, title: "Results", level: 1, isDefault: true },
        { id: 21, title: "Primary Findings", level: 2, isDefault: true },
        { id: 22, title: "Secondary Findings", level: 2, isDefault: true },
        { id: 23, title: "Discussion", level: 1, isDefault: true },
        { id: 24, title: "Interpretation of Results", level: 2, isDefault: true },
        { id: 25, title: "Limitations", level: 2, isDefault: true },
        { id: 26, title: "Implications", level: 2, isDefault: true },
        { id: 27, title: "Conclusion", level: 1, isDefault: true },
        { id: 28, title: "Summary", level: 2, isDefault: true },
        { id: 29, title: "Future Research", level: 2, isDefault: true },
        { id: 30, title: "References", level: 1, isDefault: true },
        { id: 31, title: "Appendices", level: 1, isDefault: true },
      ]

      // Create planning items from template TOC
      const templatePlanningItems = templateTocItems.map((item) => ({
        ...item,
        status: "empty",
        comments: getDefaultComment(item.title), // Add default comments based on section
        words: 0,
        paragraphs: 0,
        tables: 0,
        graphics: 0,
      }))

      setTocItems(templateTocItems)
      setPlanningItems(templatePlanningItems)
      setNextId(32) // Next ID after the template items

      // Refresh statistics for the template items
      setTimeout(() => refreshStatistics(), 500)

      // Save the template to document properties
      setTimeout(() => saveToDocumentProperties(), 1000)
    } catch (error) {
      console.error("Error creating template structure:", error)
      setError("Failed to create template structure. Please try again.")
    }
  }

  // Helper function to generate default comments based on section title
  const getDefaultComment = (title) => {
    if (!title) return ""

    const commentTemplates = {
      "Title Page": "Include title, author name, date, and institutional affiliation.",
      Abstract: "Brief summary of the entire document (150-250 words).",
      Introduction: "Introduce the topic and provide context for the reader.",
      Background: "Provide relevant background information on the topic.",
      "Problem Statement": "Clearly state the problem being addressed.",
      "Research Questions": "List the specific questions this document aims to answer.",
      "Literature Review": "Analyze and synthesize relevant existing research.",
      Methodology: "Describe the methods used to collect and analyze data.",
      Results: "Present findings without interpretation.",
      Discussion: "Interpret results and connect to existing literature.",
      Conclusion: "Summarize key findings and their implications.",
      References: "List all sources cited in the document.",
    }

    return commentTemplates[title] || ""
  }

  // Save data to document properties
  const saveToDocumentProperties = async () => {
    try {
      // Only save the necessary data (not statistics)
      const dataToSave = {
        tocItems,
        planningItems: planningItems.map((item) => ({
          id: item.id,
          title: item.title,
          level: item.level,
          status: item.status,
          comments: item.comments,
          isDefault: item.isDefault,
        })),
      }

      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
        // Save to localStorage for development
        localStorage.setItem("documentPlannerData", JSON.stringify(dataToSave))
        return
      }

      await Word.run(async (context) => {
        try {
          // Get document properties
          const properties = context.document.properties.customProperties

          // Set our data property
          properties.add("documentPlannerData", JSON.stringify(dataToSave))

          await context.sync()
          console.log("Data saved successfully")
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)
        }
      })
    } catch (error) {
      console.error("Error saving data:", error)
      setError("Failed to save data. Please try again.")
    }
  }

  // Delete all saved data
  const deleteAllData = async () => {
    try {
      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
        // Clear localStorage for development
        localStorage.removeItem("documentPlannerData")

        // Reset the state
        setTocItems([])
        setPlanningItems([])
        setNextId(1)
        setTemplateApplied(false)
        setDataLoaded(false) // This will trigger the loading process again
        setDeleteConfirmOpen(false)
        return
      }

      await Word.run(async (context) => {
        try {
          // Get document properties
          const properties = context.document.properties.customProperties

          // Delete our data property
          properties.delete("documentPlannerData")

          await context.sync()
          console.log("Data deleted successfully")

          // Reset the state
          setTocItems([])
          setPlanningItems([])
          setNextId(1)
          setTemplateApplied(false)
          setDataLoaded(false) // This will trigger the loading process again
          setDeleteConfirmOpen(false)
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)
        }
      })
    } catch (error) {
      console.error("Error deleting data:", error)
      setError("Failed to delete data. Please try again.")
    }
  }

  // Get actual statistics from the Word document
  const refreshStatistics = async () => {
    setRefreshing(true)

    try {
      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
        // Simulate statistics for development
        const updatedItems = planningItems.map((item) => ({
          ...item,
          words: item.status !== "empty" ? Math.floor(Math.random() * 500) + 50 : 0,
          paragraphs: item.status !== "empty" ? Math.floor(Math.random() * 10) + 1 : 0,
          tables: Math.random() > 0.7 ? Math.floor(Math.random() * 3) : 0,
          graphics: Math.random() > 0.8 ? Math.floor(Math.random() * 2) : 0,
        }))

        setPlanningItems(updatedItems)
        setRefreshing(false)
        return
      }

      await Word.run(async (context) => {
        try {
          // Get all headings in the document
          const body = context.document.body
          body.load("paragraphs")
          await context.sync()

          const paragraphs = body.paragraphs.items
          const headings = []

          // Identify headings and their positions
          for (let i = 0; i < paragraphs.length; i++) {
            paragraphs[i].load("text, style")
            await context.sync()

            const style = paragraphs[i].style
            if (style && (style.includes("Heading") || style === "Title")) {
              headings.push({
                text: paragraphs[i].text.trim(),
                index: i,
              })
            }
          }

          // Create sections based on headings
          const sections = []
          for (let i = 0; i < headings.length; i++) {
            const startIndex = headings[i].index + 1
            const endIndex = i < headings.length - 1 ? headings[i + 1].index : paragraphs.length

            // Count content between headings
            let wordCount = 0
            let paragraphCount = 0
            let tableCount = 0
            let graphicCount = 0

            for (let j = startIndex; j < endIndex; j++) {
              paragraphs[j].load("text")
              await context.sync()

              const text = paragraphs[j].text.trim()
              if (text.length > 0) {
                wordCount += text.split(/\s+/).length
                paragraphCount++

                // In a real implementation, you would check for tables and graphics
                // This is a simplified version
                if (text.toLowerCase().includes("table")) {
                  tableCount++
                }
                if (text.toLowerCase().includes("figure") || text.toLowerCase().includes("image")) {
                  graphicCount++
                }
              }
            }

            sections.push({
              title: headings[i].text,
              words: wordCount,
              paragraphs: paragraphCount,
              tables: tableCount,
              graphics: graphicCount,
            })
          }

          // Update planning items with actual statistics
          const updatedItems = planningItems.map((item) => {
            const matchingSection = sections.find(
              (section) =>
                section.title.toLowerCase().includes(item.title.toLowerCase()) ||
                item.title.toLowerCase().includes(section.title.toLowerCase()),
            )

            if (matchingSection) {
              return {
                ...item,
                words: matchingSection.words,
                paragraphs: matchingSection.paragraphs,
                tables: matchingSection.tables,
                graphics: matchingSection.graphics,
              }
            }

            // If no matching section found, keep existing stats or set to 0
            return {
              ...item,
              words: item.words || 0,
              paragraphs: item.paragraphs || 0,
              tables: item.tables || 0,
              graphics: item.graphics || 0,
            }
          })

          setPlanningItems(updatedItems)
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)

          // Fallback to simulated statistics
          const updatedItems = planningItems.map((item) => ({
            ...item,
            words: item.status !== "empty" ? Math.floor(Math.random() * 500) + 50 : 0,
            paragraphs: item.status !== "empty" ? Math.floor(Math.random() * 10) + 1 : 0,
            tables: Math.random() > 0.7 ? Math.floor(Math.random() * 3) : 0,
            graphics: Math.random() > 0.8 ? Math.floor(Math.random() * 2) : 0,
          }))

          setPlanningItems(updatedItems)
        }
      })
    } catch (error) {
      console.error("Error refreshing statistics:", error)
      setError("Failed to refresh statistics. Please try again.")

      // Fallback to simulated statistics
      const updatedItems = planningItems.map((item) => ({
        ...item,
        words: item.status !== "empty" ? Math.floor(Math.random() * 500) + 50 : 0,
        paragraphs: item.status !== "empty" ? Math.floor(Math.random() * 10) + 1 : 0,
        tables: Math.random() > 0.7 ? Math.floor(Math.random() * 3) : 0,
        graphics: Math.random() > 0.8 ? Math.floor(Math.random() * 2) : 0,
      }))

      setPlanningItems(updatedItems)
    } finally {
      setRefreshing(false)
    }
  }

  // Update status of a section
  const updateStatus = (id, status) => {
    try {
      setPlanningItems((prev) => prev.map((item) => (item.id === id ? { ...item, status } : item)))

      // Save after update
      setTimeout(() => saveToDocumentProperties(), 100)
    } catch (error) {
      console.error("Error updating status:", error)
      setError("Failed to update status. Please try again.")
    }
  }

  // Update comments for a section
  const updateComments = (id, comments) => {
    try {
      setPlanningItems((prev) => prev.map((item) => (item.id === id ? { ...item, comments } : item)))
      setCommentItem(null)

      // Save after update
      setTimeout(() => saveToDocumentProperties(), 100)
    } catch (error) {
      console.error("Error updating comments:", error)
      setError("Failed to update comments. Please try again.")
    }
  }

  // Calculate overall document completion
  const calculateCompletion = () => {
    try {
      const totalSections = planningItems.length
      if (totalSections === 0) return 0

      // Sum up progress for each section based on its status
      const totalProgress = planningItems.reduce((acc, item) => acc + (STATUS_PROGRESS[item.status] || 0), 0)

      // Normalize to percentage scale
      return (totalProgress / (totalSections * 100)) * 100
    } catch (error) {
      console.error("Error calculating completion:", error)
      return 0
    }
  }

  // Sync plan with document headers
  const syncPlanWithDocument = async () => {
    try {
      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available");
        alert("This feature requires the Word API, which is not available in this environment.");
        return;
      }

      await Word.run(async (context) => {
        try {
          // Get all headings in the document
          const body = context.document.body;
          body.load("paragraphs");
          await context.sync();

          const paragraphs = body.paragraphs.items;
          const documentHeadings = [];
          
          // Identify headings and their positions
          for (let i = 0; i < paragraphs.length; i++) {
            paragraphs[i].load("text, style");
            await context.sync();
          
            const style = paragraphs[i].style;
            if (style && (style.includes("Heading1") || style === "Title")) {
              documentHeadings.push({
                text: paragraphs[i].text.trim(),
                level: 1,
                index: i
              });
            } else if (style && style.includes("Heading2")) {
              documentHeadings.push({
                text: paragraphs[i].text.trim(),
                level: 2,
                index: i
              });
            }
          }

          // If no headings found in document
          if (documentHeadings.length === 0) {
            alert("No headings found in the document. Nothing to sync.");
            return;
          }

          // Find headings in document but not in plan
          const headingsToAddToPlan = [];
          for (const docHeading of documentHeadings) {
            const exists = planningItems.some(
              item => item.title.toLowerCase() === docHeading.text.toLowerCase()
            );
          
            if (!exists) {
              headingsToAddToPlan.push({
                title: docHeading.text,
                level: docHeading.level,
                index: docHeading.index
              });
            }
          }

          // Find headings in plan but not in document (only L1 items)
          const headingsToAddToDocument = planningItems.filter(
            item => 
              item.level === 1 && 
              !documentHeadings.some(dh => dh.text.toLowerCase() === item.title.toLowerCase())
          );

          // Add new headings to plan
          if (headingsToAddToPlan.length > 0) {
            const newPlanningItems = [...planningItems];
            const newTocItems = [...tocItems];
          
            for (const heading of headingsToAddToPlan) {
              const id = nextId + headingsToAddToPlan.indexOf(heading);
            
              // Find the right position to insert based on document order
              let insertIndex = newPlanningItems.length;
              for (let i = 0; i < newPlanningItems.length; i++) {
                const matchingDocHeading = documentHeadings.find(
                  dh => dh.text.toLowerCase() === newPlanningItems[i].title.toLowerCase()
                );
              
                if (matchingDocHeading && matchingDocHeading.index > heading.index) {
                  insertIndex = i;
                  break;
                }
              }
            
              const newItem = {
                id,
                title: heading.title,
                level: heading.level,
                status: "created", // Mark as created since it exists in document
                comments: getDefaultComment(heading.title),
                words: 0,
                paragraphs: 0,
                tables: 0,
                graphics: 0,
                isDefault: false,
              };
            
            // Insert at the right position
              newPlanningItems.splice(insertIndex, 0, newItem);
              newTocItems.splice(insertIndex, 0, {
                id,
                title: heading.title,
                level: heading.level,
                isDefault: false
              });
            }
          
            setPlanningItems(newPlanningItems);
            setTocItems(newTocItems);
            setNextId(nextId + headingsToAddToPlan.length);
          }

          // Add missing L1 headings to document
          if (headingsToAddToDocument.length > 0) {
            for (const item of headingsToAddToDocument) {
              // Insert heading
              const paragraph = context.document.body.insertParagraph(item.title, "End");
            
              // Set formatting
              if (paragraph && paragraph.font) {
                paragraph.font.set({
                  size: 16,
                  bold: true,
                });
                if (Word.Style && Word.Style.heading1) {
                  paragraph.styleBuiltIn = Word.Style.heading1;
                }
              }
            
              // Insert template text
              const templateText = context.document.body.insertParagraph("<insert your text here>", "End");
              if (templateText && templateText.font) {
                templateText.font.set({
                  italic: true,
                  size: 11,
                  color: "#666666"
                });
              }
            
              // Insert a paragraph break
              context.document.body.insertParagraph("", "End");
            }
          }

          await context.sync();
        
          // Save changes to document properties
          setTimeout(() => saveToDocumentProperties(), 100);
        
          // Refresh statistics
          setTimeout(() => refreshStatistics(), 500);
        
          // Show summary
          const message = `Sync complete!\n\n${headingsToAddToPlan.length} headings added to plan.\n${headingsToAddToDocument.length} headings added to document.`;
          alert(message);
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError);
          alert("Failed to sync plan with document. Please try again.");
        }
      });
    } catch (error) {
      console.error("Error syncing plan with document:", error);
      setError("Failed to sync plan with document. Please try again.");
    }
  };

  // Update section title
  const updateTitle = (id, title) => {
    try {
      setPlanningItems((prev) => prev.map((item) => (item.id === id ? { ...item, title } : item)))

      // Also update in TOC items
      setTocItems((prev) => prev.map((item) => (item.id === id ? { ...item, title } : item)))

      setEditingItem(null)

      // Save after update
      setTimeout(() => saveToDocumentProperties(), 100)
    } catch (error) {
      console.error("Error updating title:", error)
      setError("Failed to update title. Please try again.")
    }
  }

  // Add a new section
  const addSection = (level = 1) => {
    try {
      const newItem = {
        id: nextId,
        title: "New Section",
        level,
        status: "empty",
        comments: "",
        words: 0,
        paragraphs: 0,
        tables: 0,
        graphics: 0,
        isDefault: false,
      }

      setPlanningItems((prev) => [...prev, newItem])
      setTocItems((prev) => [...prev, { id: nextId, title: "New Section", level, isDefault: false }])
      setNextId((prev) => prev + 1)
      setEditingItem(newItem.id)

      // Save after update
      setTimeout(() => saveToDocumentProperties(), 100)
    } catch (error) {
      console.error("Error adding section:", error)
      setError("Failed to add section. Please try again.")
    }
  }

  // Delete a section
  const deleteSection = (id) => {
    try {
      // Check if the item is a default item that shouldn't be deleted
      const itemToDelete = planningItems.find((item) => item.id === id)
      if (itemToDelete && itemToDelete.isDefault) {
        alert("Default sections cannot be deleted.")
        return
      }

      setPlanningItems((prev) => prev.filter((item) => item.id !== id))
      setTocItems((prev) => prev.filter((item) => item.id !== id))

      // Save after update
      setTimeout(() => saveToDocumentProperties(), 100)
    } catch (error) {
      console.error("Error deleting section:", error)
      setError("Failed to delete section. Please try again.")
    }
  }

  // Create TOC scaffold in the document
  const createTocScaffold = async () => {
    setBuildingToc(true);

    try {
      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available");
        alert("This feature requires the Word API, which is not available in this environment.");
        setBuildingToc(false);
        return;
      }

    await Word.run(async (context) => {
      try {
        // First sync plan with document to ensure everything is up to date
        await syncPlanWithDocument();
        
        // Sort items by ID to maintain the correct order
        const sortedItems = [...tocItems].sort((a, b) => a.id - b.id);

        // Insert a title for the TOC
        const titleParagraph = context.document.body.insertParagraph("TABLE OF CONTENTS", "Start");
        if (titleParagraph && titleParagraph.font) {
          titleParagraph.font.set({
            bold: true,
            size: 16,
          });
        }

        // Insert each TOC item with appropriate indentation
        for (const item of sortedItems) {
          const indent = "  ".repeat(item.level - 1);
          const paragraph = context.document.body.insertParagraph(`${indent}${item.title}`, "End");

          // Set indentation based on level
          if (paragraph) {
            paragraph.leftIndent = (item.level - 1) * 20;
          }
          // Set formatting based on level
          if (paragraph && paragraph.font) {
            if (item.level === 1) {
              paragraph.font.set({
                size: 16,
                bold: true,
              })
              if (Word.Style && Word.Style.heading1) {
                paragraph.styleBuiltIn = Word.Style.heading1
              }
            } else {
              paragraph.font.set({
                size: 14,
                bold: true,
              })
              if (Word.Style && Word.Style.heading2) {
                paragraph.styleBuiltIn = Word.Style.heading2
              }
            }
          }
        }

        await context.sync();
        alert("TOC scaffold has been created in the document!");
      } catch (contextError) {
        console.error("Error in Word.run context:", contextError);
        alert("Failed to create TOC scaffold. Please try again.");
      }
    });
  } catch (error) {
    console.error("Error creating TOC scaffold:", error);
    alert("Failed to create TOC scaffold. Please try again.");
    setError("Failed to create TOC scaffold. Please try again.");
  } finally {
    setBuildingToc(false);
  }
};

  // Build document structure with headers
  const buildDocumentStructure = async () => {
    setBuildingDocument(true)

    try {
      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
        alert("This feature requires the Word API, which is not available in this environment.")
        setBuildingDocument(false)
        return
      }

      await Word.run(async (context) => {
        try {
          // Sort items by ID to maintain the correct order
          const sortedItems = [...planningItems].sort((a, b) => a.id - b.id)

          // Insert each item as a header with appropriate formatting
          for (const item of sortedItems) {
            const paragraph = context.document.body.insertParagraph(item.title, "End")

            // Set formatting based on level
            if (paragraph && paragraph.font) {
              if (item.level === 1) {
                paragraph.font.set({
                  size: 16,
                  bold: true,
                })
                if (Word.Style && Word.Style.heading1) {
                  paragraph.styleBuiltIn = Word.Style.heading1
                }
              } else {
                paragraph.font.set({
                  size: 14,
                  bold: true,
                })
                if (Word.Style && Word.Style.heading2) {
                  paragraph.styleBuiltIn = Word.Style.heading2
                }
              }
            }

            // Insert template text under each header
            const templateText = context.document.body.insertParagraph("<insert your text here>", "End")
            if (templateText && templateText.font) {
              templateText.font.set({
                italic: true,
                size: 11,
                color: "#666666",
              })
            }

            // Insert a paragraph break after each section
            context.document.body.insertParagraph("", "End")
          }

          await context.sync()
          alert("Document structure has been built with headers!")
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)
          alert("Failed to build document structure. Please try again.")
        }
      })
    } catch (error) {
      console.error("Error building document structure:", error)
      alert("Failed to build document structure. Please try again.")
      setError("Failed to build document structure. Please try again.")
    } finally {
      setBuildingDocument(false)
    }
  }

  // Custom render function for planning items
  const renderPlanningItem = (item) => {
    return (
      <div key={item.id} style={{ marginBottom: "8px", borderBottom: "1px solid #f0f0f0", paddingBottom: "4px" }}>
        {/* First row: Section name and status */}
        <div className={sectionRowStyle} style={{ paddingLeft: (item.level - 1) * 10 }}>
          {editingItem === item.id ? (
            <TextField
              defaultValue={item.title}
              autoFocus
              onBlur={(e) => {
                if (e && e.target) {
                  updateTitle(item.id, e.target.value)
                }
              }}
              onKeyDown={(e) => {
                if (e && e.key === "Enter" && e.target) {
                  updateTitle(item.id, e.target.value)
                }
              }}
              styles={{ root: { width: "70%", minWidth: 80 } }}
            />
          ) : (
            <span
              className={sectionNameStyle}
              onClick={() => !item.isDefault && setEditingItem(item.id)}
              title={item.isDefault ? "Default sections cannot be edited" : "Click to edit"}
              style={{
                width: "70%",
                overflow: "hidden",
                textOverflow: "ellipsis",
                cursor: item.isDefault ? "default" : "pointer",
                textDecoration: item.isDefault ? "none" : undefined,
              }}
            >
              {item.title}
            </span>
          )}
          <Dropdown
            selectedKey={item.status}
            options={STATUS_OPTIONS}
            onChange={(e, option) => {
              if (option) {
                updateStatus(item.id, option.key)
              }
            }}
            styles={{
              dropdown: {
                width: "30%",
                minWidth: 70,
                maxWidth: 90,
              },
              title: {
                backgroundColor: STATUS_COLORS[item.status]?.background,
                color: STATUS_COLORS[item.status]?.color,
                borderColor: "transparent",
                fontSize: "11px",
                padding: "0 4px",
              },
              caretDown: {
                fontSize: "8px",
              },
            }}
          />
        </div>

        {/* Second row: Action buttons */}
        <div className={buttonRowStyle}>
          <TooltipHost content="Edit Comments">
            <IconButton
              iconProps={{ iconName: "Comment" }}
              onClick={() => setCommentItem(item.id)}
              styles={{ root: { height: 24, width: 24, marginRight: 8 } }}
            />
          </TooltipHost>

          <TooltipHost content="View Statistics">
            <IconButton
              iconProps={{ iconName: "BarChart4" }}
              onClick={(e) => {
                if (e && e.currentTarget) {
                  setStatsItem(item)
                  setStatsCalloutTarget(e.currentTarget)
                  setStatsCalloutVisible(true)
                }
              }}
              styles={{ root: { height: 24, width: 24, marginRight: 8 } }}
            />
          </TooltipHost>

          <TooltipHost content={item.isDefault ? "Default sections cannot be deleted" : "Delete Section"}>
            <IconButton
              iconProps={{ iconName: "Delete" }}
              onClick={() => deleteSection(item.id)}
              disabled={item.isDefault}
              styles={{ root: { height: 24, width: 24 } }}
            />
          </TooltipHost>
        </div>
      </div>
    )
  }

  // If Office is not initialized
  if (!isOfficeInitialized) {
    return (
      <Stack styles={containerStyles}>
        <Spinner label="Loading Office.js..." size={SpinnerSize.large} />
      </Stack>
    )
  }

  return (
    <Stack styles={containerStyles} ref={containerRef}>
      {/* Error message */}
      {error && (
        <div style={{ color: "red", marginBottom: 10, padding: 10, backgroundColor: "#fff4ce" }}>
          {error}
          <IconButton iconProps={{ iconName: "Cancel" }} onClick={() => setError(null)} style={{ float: "right" }} />
        </div>
      )}

      {/* Header */}
      <Stack horizontal horizontalAlign="space-between" styles={headerStyles}>
        <Stack>
          <h1 style={titleStyles.root}>Writing Planner</h1>
          <p style={subtitleStyles.root}>Plan your work and focus on your magic.</p>
        </Stack>
        <Stack horizontal tokens={{ childrenGap: 5 }}>
          <TooltipHost content="Refresh Statistics">
            <IconButton iconProps={{ iconName: "Refresh" }} onClick={refreshStatistics} disabled={refreshing} />
          </TooltipHost>
          <TooltipHost content="Delete My Data">
            <IconButton iconProps={{ iconName: "Delete" }} onClick={() => setDeleteConfirmOpen(true)} />
          </TooltipHost>
          <TooltipHost content="About">
            <IconButton iconProps={{ iconName: "Info" }} onClick={() => setAboutOpen(true)} />
          </TooltipHost>
        </Stack>
      </Stack>

      {/* Progress */}
      <Stack tokens={{ childrenGap: 10, padding: "10px 0" }}>
        <Stack horizontal horizontalAlign="space-between">
          <Label>Progress:</Label>
          <Label>{Math.round(calculateCompletion())}%</Label>
        </Stack>
        <ProgressIndicator percentComplete={calculateCompletion() / 100} />
      </Stack>

      {/* Action Buttons */}
      <Stack horizontal tokens={{ childrenGap: 10, padding: "10px 0" }}>
        <DefaultButton text="Sync" iconProps={{ iconName: "Sync" }} onClick={syncPlanWithDocument} />
        <PrimaryButton text="Add Section" iconProps={{ iconName: "Add" }} onClick={() => addSection()} />
      </Stack>

      {/* Tabs */}
      <Pivot
        selectedKey={activeTab}
        onLinkClick={(item) => item && setActiveTab(item.props.itemKey)}
        styles={{ root: { marginBottom: 10 } }}
      >
        <PivotItem headerText="Planning" itemKey="plan" itemIcon="FileDocument">
          <Stack tokens={{ childrenGap: 10 }}>
            <div style={{ overflowX: "hidden", maxHeight: "400px", overflowY: "auto" }}>
              {planningItems.map(renderPlanningItem)}
            </div>
            <PrimaryButton
              text={buildingDocument ? "Building..." : "Build Document Structure"}
              iconProps={{ iconName: "BuildDefinition" }}
              onClick={buildDocumentStructure}
              disabled={buildingDocument}
            />
          </Stack>
        </PivotItem>
        <PivotItem headerText="TOC Template" itemKey="toc" itemIcon="BulletedList">
          <Stack tokens={{ childrenGap: 10 }}>
            <div style={{ maxHeight: 200, overflowY: "auto" }}>
              {tocItems.map((item) => (
                <div
                  key={item.id}
                  style={{
                    paddingLeft: (item.level - 1) * 15,
                    marginBottom: 5,
                    display: "flex",
                    alignItems: "center",
                  }}
                >
                  <span>{item.title}</span>
                  {planningItems.find((p) => p.id === item.id)?.status === "empty" && (
                    <span style={{ marginLeft: 5, color: "#c19c00" }}>⚠</span>
                  )}
                  {item.isDefault && <span style={{ marginLeft: 5, fontSize: "10px", color: "#666" }}>(default)</span>}
                </div>
              ))}
            </div>
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <DefaultButton
                text="Add L1"
                iconProps={{ iconName: "Add" }}
                onClick={() => addSection(1)}
                styles={{ root: { flexGrow: 1 } }}
              />
              <DefaultButton
                text="Add L2"
                iconProps={{ iconName: "Add" }}
                onClick={() => addSection(2)}
                styles={{ root: { flexGrow: 1 } }}
              />
            </Stack>
            <PrimaryButton
              text={buildingToc ? "Creating..." : "Create TOC in Document"}
              iconProps={{ iconName: "FileTemplate" }}
              onClick={createTocScaffold}
              disabled={buildingToc}
            />
          </Stack>
        </PivotItem>
      </Pivot>

      {/* Statistics Callout */}
      {statsCalloutVisible && statsCalloutTarget && statsItem && (
        <Callout target={statsCalloutTarget} onDismiss={() => setStatsCalloutVisible(false)} setInitialFocus>
          <Stack tokens={{ padding: 20, childrenGap: 10 }}>
            <Label>Section Statistics</Label>
            <Stack tokens={{ childrenGap: 5 }}>
              <Stack horizontal horizontalAlign="space-between">
                <span>Words:</span>
                <strong>{statsItem.words || "0"}</strong>
              </Stack>
              <Stack horizontal horizontalAlign="space-between">
                <span>Paragraphs:</span>
                <strong>{statsItem.paragraphs || "0"}</strong>
              </Stack>
              <Stack horizontal horizontalAlign="space-between">
                <span>Tables:</span>
                <strong>{statsItem.tables || "0"}</strong>
              </Stack>
              <Stack horizontal horizontalAlign="space-between">
                <span>Graphics:</span>
                <strong>{statsItem.graphics || "0"}</strong>
              </Stack>
            </Stack>
          </Stack>
        </Callout>
      )}

      {/* Comments Panel */}
      <Panel
        isOpen={commentItem !== null}
        onDismiss={() => setCommentItem(null)}
        headerText="Section Comments"
        closeButtonAriaLabel="Close"
      >
        {commentItem !== null && (
          <Stack tokens={{ childrenGap: 15, padding: "20px 0" }}>
            <TextField
              label="Comments"
              multiline
              rows={5}
              value={planningItems.find((item) => item.id === commentItem)?.comments || ""}
              onChange={(e, newValue) => {
                setPlanningItems((prev) =>
                  prev.map((item) => (item.id === commentItem ? { ...item, comments: newValue || "" } : item)),
                )
              }}
            />
            <PrimaryButton
              text="Save Comments"
              onClick={() => {
                const item = planningItems.find((item) => item.id === commentItem)
                if (item) {
                  updateComments(item.id, item.comments)
                }
              }}
            />
          </Stack>
        )}
      </Panel>

      {/* About Dialog */}
      <Dialog
        hidden={!aboutOpen}
        onDismiss={() => setAboutOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "About Writing Planner",
          subText: "Created with joy by V0.dev (CC)2025",
        }}
      >
        <div style={{ margin: "20px 0" }}>
          <p>Contact: ali.vakilzadeh@gmail.com</p>
          <p style={{ marginTop: 10 }}>This add-in helps you plan and structure your documents before writing.</p>
        </div>
        <DialogFooter>
          <PrimaryButton text="Close" onClick={() => setAboutOpen(false)} />
        </DialogFooter>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <Dialog
        hidden={!deleteConfirmOpen}
        onDismiss={() => setDeleteConfirmOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Confirm Deletion",
          subText: "All your plan data will be lost! Are you sure?",
        }}
      >
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={() => setDeleteConfirmOpen(false)} />
          <PrimaryButton text="Yes, Delete Everything" onClick={deleteAllData} />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
}

