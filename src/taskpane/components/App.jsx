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

  // Add this to the state variables near the top of the component (around line 120)
  const [showTemplateButton, setShowTemplateButton] = useState(true)
  const [deleteButtonTimer, setDeleteButtonTimer] = useState(null)
  const [selectedTemplateItems, setSelectedTemplateItems] = useState({})
  const [editingTocItem, setEditingTocItem] = useState(null)
  // Add a new state variable after the other state variables (around line 120)
  const [templateJustApplied, setTemplateJustApplied] = useState(false)

  // make sure the planningItems are saved to localStorage whenever they change
  useEffect(() => {
    localStorage.setItem("planningItems", JSON.stringify(planningItems))
  }, [planningItems])

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

  // Replace the existing useEffect for template application with this:
  useEffect(() => {
    if (dataLoaded && tocItems.length === 0 && planningItems.length === 0) {
      // Don't automatically apply template, just show the template button
      setShowTemplateButton(true)
    } else if (planningItems.length > 0) {
      // If we have planning items, hide the template button
      setShowTemplateButton(false)
    }
  }, [dataLoaded, tocItems.length, planningItems.length])

  // Check if document is empty
  const checkIfDocumentIsEmpty = async () => {
    try {
      // Skip checking if the template has already been applied
      if (templateApplied) {
        setDocumentIsEmpty(false)
        return false
      }
      if (!Word || typeof Word.run !== "function") {
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

  // Replace the existing loadFromDocumentProperties function with this improved version:
  const loadFromDocumentProperties = async () => {
    try {
      console.log("Starting to load document properties")

      // First check if document is empty
      // Skip checking if the template has already been applied
      let isEmpty = false
      if (templateApplied) {
        setDocumentIsEmpty(false)
      } else {
        isEmpty = await checkIfDocumentIsEmpty()
      }

      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.log("Word API not available, trying localStorage")

        // Try to get from localStorage for development
        const plannerData = localStorage.getItem("documentPlannerData")

        if (plannerData) {
          try {
            console.log("Found data in localStorage:", plannerData.substring(0, 100) + "...")
            const data = JSON.parse(plannerData)

            if (data && data.tocItems && data.planningItems) {
              console.log(`Loading ${data.tocItems.length} TOC items and ${data.planningItems.length} planning items`)

              // Set template applied to true to prevent overwriting with template
              setTemplateApplied(true)

              setTocItems(data.tocItems || [])
              setPlanningItems(
                (data.planningItems || []).map((item) => ({
                  ...item,
                  words: item.words || 0,
                  paragraphs: item.paragraphs || 0,
                  tables: item.tables || 0,
                  graphics: item.graphics || 0,
                })),
              )

              // Find the highest ID to set nextId correctly
              const highestId = Math.max(...data.planningItems.map((item) => item.id || 0), 0)
              setNextId(highestId + 1)

              // Initialize selected items for template items
              const initialSelectedItems = {}
              data.tocItems.forEach((item) => {
                if (item.isDefault) {
                  initialSelectedItems[item.id] = true
                }
              })
              setSelectedTemplateItems(initialSelectedItems)

              // Refresh statistics after loading data
              setTimeout(() => refreshStatistics(), 500)
            }
          } catch (parseError) {
            console.error("Error parsing planner data:", parseError)
          }
        } else {
          console.log("No data found in localStorage")
        }

        setDataLoaded(true)
        return
      }

      console.log("Using Word API to load data")
      await Word.run(async (context) => {
        try {
          // Get document properties
          const properties = context.document.properties.customProperties
          properties.load("items")
          await context.sync()

          console.log("Loaded document properties")

          // Find our data property
          let plannerData = null
          if (properties.items && Array.isArray(properties.items)) {
            console.log(`Found ${properties.items.length} custom properties`)

            for (let i = 0; i < properties.items.length; i++) {
              properties.items[i].load("key,value")
            }
            await context.sync()

            for (let i = 0; i < properties.items.length; i++) {
              if (properties.items[i] && properties.items[i].key === "documentPlannerData") {
                plannerData = properties.items[i].value
                console.log("Found documentPlannerData property")
                break
              }
            }
          } else {
            console.log("No custom properties found or items is not an array")
          }

          if (!plannerData) {
            // Try to get from localStorage as fallback
            plannerData = localStorage.getItem("documentPlannerData")
            console.log("Falling back to localStorage")
          }

          if (plannerData) {
            try {
              console.log("Parsing planner data")
              const data = JSON.parse(plannerData)

              if (data && data.tocItems && data.planningItems) {
                console.log(`Loading ${data.tocItems.length} TOC items and ${data.planningItems.length} planning items`)

                // Set template applied to true to prevent overwriting with template
                setTemplateApplied(true)

                // Find the highest ID to set nextId correctly
                const highestId = Math.max(...data.planningItems.map((item) => item.id || 0), 0)

                setTocItems(data.tocItems || [])

                // Set planning items but preserve statistics if they exist
                setPlanningItems(
                  (data.planningItems || []).map((item) => ({
                    ...item,
                    words: item.words || 0,
                    paragraphs: item.paragraphs || 0,
                    tables: item.tables || 0,
                    graphics: item.graphics || 0,
                  })),
                )

                setNextId(highestId + 1)

                // Initialize selected items for template items
                const initialSelectedItems = {}
                data.tocItems.forEach((item) => {
                  if (item.isDefault) {
                    initialSelectedItems[item.id] = true
                  }
                })
                setSelectedTemplateItems(initialSelectedItems)

                // Refresh statistics after loading data
                setTimeout(() => refreshStatistics(), 500)
              }
            } catch (parseError) {
              console.error("Error parsing planner data:", parseError)
            }
          } else if (isEmpty) {
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
      setShowTemplateButton(false)
      setTemplateJustApplied(true) // Set flag to prevent immediate refresh

      // Refresh statistics for the template items
      setTimeout(() => refreshStatistics(), 1000)

      // Save the template to document properties
      setTimeout(() => saveToDocumentProperties(), 1000)
      saveToDocumentProperties()
        .then(() => {
          console.log("397-Template saved successfully")
        })
        .catch((error) => {
          console.error("Error saving template:", error)
          setError("Failed to save template. Please try again.")
        })

      // Add this after setting the template items:
      const initialSelectedItems = {}
      templateTocItems.forEach((item) => {
        initialSelectedItems[item.id] = true
      })
      setSelectedTemplateItems(initialSelectedItems)
    } catch (error) {
      console.error("Error creating template structure:", error)
      setError("Failed to create template structure. Please try again.")
    }
  }

  // Also update the saveToDocumentProperties function to ensure data is saved correctly
  // Around line 450, update the saveToDocumentProperties function:

  const saveToDocumentProperties = async () => {
    try {
      console.log("Saving data to document properties")

      // Only save the necessary data (including statistics for persistence)
      const dataToSave = {
        tocItems,
        planningItems: planningItems.map((item) => ({
          id: item.id,
          title: item.title,
          level: item.level,
          status: item.status,
          comments: item.comments,
          isDefault: item.isDefault,
          words: item.words || 0,
          paragraphs: item.paragraphs || 0,
          tables: item.tables || 0,
          graphics: item.graphics || 0,
        })),
      }

      // Save to localStorage for backup and development
      localStorage.setItem("documentPlannerData", JSON.stringify(dataToSave))
      console.log("Data saved to localStorage")

      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
        return Promise.resolve() // Resolve since we saved to localStorage
      }

      await Word.run(async (context) => {
        try {
          // Get document properties
          const properties = context.document.properties.customProperties
          properties.load("items")
          await context.sync()

          // Remove existing property if it exists
          let existingPropFound = false
          if (properties.items && Array.isArray(properties.items)) {
            for (let i = 0; i < properties.items.length; i++) {
              properties.items[i].load("key")
            }
            await context.sync()

            for (let i = 0; i < properties.items.length; i++) {
              if (properties.items[i] && properties.items[i].key === "documentPlannerData") {
                properties.items[i].delete()
                existingPropFound = true
                console.log("Deleted existing documentPlannerData property")
                break
              }
            }

            if (existingPropFound) {
              await context.sync()
            }
          }

          // Set our data property
          properties.add("documentPlannerData", JSON.stringify(dataToSave))
          console.log("Added new documentPlannerData property")

          await context.sync()
          console.log("Data saved successfully to document properties")
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)
        }
      })
      return Promise.resolve()
    } catch (error) {
      console.error("Error saving data:", error)
      setError("Failed to save data. Please try again.")
      return Promise.reject(error)
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
          setError("Data deleted successfully")

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

  // Update the refreshStatistics function to correctly extract data
  // Around line 500, update the refreshStatistics function:

  // Modify the refreshStatistics function to check for templateJustApplied
  // Replace the existing refreshStatistics function with this:
  const refreshStatistics = async () => {
    setRefreshing(true)

    // If template was just applied, don't refresh from Word API
    if (templateJustApplied) {
      setTemplateJustApplied(false)
      setRefreshing(false)
      return
    }

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
          console.log("Refreshing statistics from document")

          // Get all content in the document
          const body = context.document.body
          body.load("paragraphs,tables,inlinePictures")
          await context.sync()

          const paragraphs = body.paragraphs.items
          const tables = body.tables ? body.tables.items : []
          const pictures = body.inlinePictures ? body.inlinePictures.items : []

          console.log(
            `Document has ${paragraphs.length} paragraphs, ${tables.length} tables, ${pictures.length} pictures`,
          )

          // Load paragraph properties
          for (let i = 0; i < paragraphs.length; i++) {
            paragraphs[i].load("text,style")
          }
          await context.sync()

          // Identify headings and their positions
          const headings = []
          for (let i = 0; i < paragraphs.length; i++) {
            const style = paragraphs[i].style
            const text = paragraphs[i].text.trim()

            if (style && (style.includes("Heading") || style === "Title") && text) {
              headings.push({
                text: text,
                index: i,
                level: style.includes("Heading1") || style === "Title" ? 1 : style.includes("Heading2") ? 2 : 3,
              })
            }
          }

          console.log(`Found ${headings.length} headings in document`)

          // If no headings, can't calculate section statistics
          if (headings.length === 0) {
            setRefreshing(false)
            return
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

            // Count paragraphs and words
            for (let j = startIndex; j < endIndex; j++) {
              const text = paragraphs[j].text.trim()
              if (text.length > 0) {
                wordCount += text.split(/\s+/).filter((word) => word.length > 0).length
                paragraphCount++
              }
            }

            // Count tables in this section (approximate)
            for (let t = 0; t < tables.length; t++) {
              // This is an approximation since we can't easily determine which section a table belongs to
              // We'll assign tables to sections based on their position relative to headings
              if (t >= startIndex && t < endIndex) {
                tableCount++
              }
            }

            // Count pictures in this section (approximate)
            for (let p = 0; p < pictures.length; p++) {
              // Similar approximation for pictures
              if (p >= startIndex && p < endIndex) {
                graphicCount++
              }
            }

            sections.push({
              title: headings[i].text,
              level: headings[i].level,
              words: wordCount,
              paragraphs: paragraphCount,
              tables: tableCount,
              graphics: graphicCount,
            })
          }

          console.log(`Created ${sections.length} sections with statistics`)

          // Update planning items with actual statistics
          const updatedItems = planningItems.map((item) => {
            // Try to find an exact match first
            let matchingSection = sections.find((section) => section.title.toLowerCase() === item.title.toLowerCase())

            // If no exact match, try a partial match
            if (!matchingSection) {
              matchingSection = sections.find(
                (section) =>
                  section.title.toLowerCase().includes(item.title.toLowerCase()) ||
                  item.title.toLowerCase().includes(section.title.toLowerCase()),
              )
            }

            if (matchingSection) {
              return {
                ...item,
                words: matchingSection.words,
                paragraphs: matchingSection.paragraphs,
                tables: matchingSection.tables,
                graphics: matchingSection.graphics,
                // Update status if it has content but is still marked as empty
                status: item.status === "empty" && matchingSection.words > 0 ? "created" : item.status,
              }
            }

            // If no matching section found, keep existing stats
            return item
          })

          setPlanningItems(updatedItems)
          console.log("Statistics updated successfully")
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

      //setPlanningItems(updatedItems)
    } finally {
      setRefreshing(false)
    }
  }

  // Update status of a section
  const updateStatus = (id, status) => {
    try {
      setPlanningItems((prev) => prev.map((item) => (item.id === id ? { ...item, status } : item)))
      // Save after update
      setTimeout(() => saveToDocumentProperties(), 1000)
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
      setTimeout(() => saveToDocumentProperties(), 1000)
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

  // Update the syncPlanWithDocument function to maintain document order
  // Around line 600, update the syncPlanWithDocument function:

  const syncPlanWithDocument = async () => {
    try {
      setError("Syncing plan with document...")

      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
        setError("This feature requires the Word API, which is not available in this environment.")
        return
      }

      await Word.run(async (context) => {
        try {
          // Get all headings in the document
          const body = context.document.body
          body.load("paragraphs")
          await context.sync()

          const paragraphs = body.paragraphs.items
          const documentHeadings = []

          // Identify headings and their positions
          for (let i = 0; i < paragraphs.length; i++) {
            paragraphs[i].load("text, style")
            await context.sync()

            const style = paragraphs[i].style
            if (style && (style.includes("Heading1") || style === "Title")) {
              documentHeadings.push({
                text: paragraphs[i].text.trim(),
                level: 1,
                index: i,
              })
            } else if (style && style.includes("Heading2")) {
              documentHeadings.push({
                text: paragraphs[i].text.trim(),
                level: 2,
                index: i,
              })
            }
          }

          // If no headings found in document
          if (documentHeadings.length === 0) {
            setError("No headings found in the document. Nothing to sync.")
            return
          }

          console.log(`Found ${documentHeadings.length} headings in document`)

          // Find headings in document but not in plan
          const headingsToAddToPlan = []
          for (const docHeading of documentHeadings) {
            const exists = planningItems.some((item) => item.title.toLowerCase() === docHeading.text.toLowerCase())

            if (!exists) {
              headingsToAddToPlan.push({
                title: docHeading.text,
                level: docHeading.level,
                index: docHeading.index,
              })
            }
          }

          // Find headings in plan but not in document (only L1 items)
          const headingsToAddToDocument = planningItems.filter(
            (item) =>
              item.level === 1 && !documentHeadings.some((dh) => dh.text.toLowerCase() === item.title.toLowerCase()),
          )

          // Create a new array for reorganized planning items
          const newPlanningItems = []
          const newTocItems = []

          // First, add items that are in the document in the order they appear
          for (const docHeading of documentHeadings) {
            // Find matching item in planning items
            const existingItem = planningItems.find(
              (item) => item.title.toLowerCase() === docHeading.text.toLowerCase(),
            )

            if (existingItem) {
              // Add existing item to the new array
              newPlanningItems.push(existingItem)

              // Also add to new TOC items
              const existingTocItem = tocItems.find((item) => item.id === existingItem.id)
              if (existingTocItem) {
                newTocItems.push(existingTocItem)
              }
            }
          }

          // Then add items that are in plan but not in document
          for (const planItem of planningItems) {
            const exists = documentHeadings.some((dh) => dh.text.toLowerCase() === planItem.title.toLowerCase())

            if (!exists) {
              // Add to new arrays if not already added
              if (!newPlanningItems.some((item) => item.id === planItem.id)) {
                newPlanningItems.push(planItem)

                // Also add to new TOC items
                const existingTocItem = tocItems.find((item) => item.id === planItem.id)
                if (existingTocItem && !newTocItems.some((item) => item.id === existingTocItem.id)) {
                  newTocItems.push(existingTocItem)
                }
              }
            }
          }

          // Add new headings to plan
          if (headingsToAddToPlan.length > 0) {
            let currentNextId = nextId

            for (const heading of headingsToAddToPlan) {
              const id = currentNextId++

              // Find the right position to insert based on document order
              let insertIndex = newPlanningItems.length
              for (let i = 0; i < newPlanningItems.length; i++) {
                const matchingDocHeading = documentHeadings.find(
                  (dh) => dh.text.toLowerCase() === newPlanningItems[i].title.toLowerCase(),
                )

                if (matchingDocHeading && matchingDocHeading.index > heading.index) {
                  insertIndex = i
                  break
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
              }

              // Insert at the right position
              newPlanningItems.splice(insertIndex, 0, newItem)
              newTocItems.splice(insertIndex, 0, {
                id,
                title: heading.title,
                level: heading.level,
                isDefault: false,
              })
            }

            setNextId(currentNextId)
          }

          // Update state with reorganized items
          setPlanningItems(newPlanningItems)
          setTocItems(newTocItems)

          // Add missing L1 headings to document
          if (headingsToAddToDocument.length > 0) {
            for (const item of headingsToAddToDocument) {
              // Insert heading
              const paragraph = context.document.body.insertParagraph(item.title, "End")

              // Set formatting
              if (paragraph && paragraph.font) {
                paragraph.font.set({
                  size: 16,
                  bold: true,
                })
                if (Word.Style && Word.Style.heading1) {
                  paragraph.styleBuiltIn = Word.Style.heading1
                }
              }

              // Insert template text
              const templateText = context.document.body.insertParagraph("<insert your text here>", "End")
              if (templateText && templateText.font) {
                templateText.font.set({
                  italic: true,
                  size: 11,
                  color: "#666666",
                })
              }

              // Insert a paragraph break
              context.document.body.insertParagraph("", "End")
            }
          }

          await context.sync()

          // Save changes to document properties
          await saveToDocumentProperties()

          // Refresh statistics
          setTimeout(() => refreshStatistics(), 500)

          // Show summary
          setError(
            `Sync complete! ${headingsToAddToPlan.length} headings added to plan. ${headingsToAddToDocument.length} headings added to document.`,
          )
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)
          setError("Failed to sync plan with document. Please try again.")
        }
      })
    } catch (error) {
      console.error("Error syncing plan with document:", error)
      setError("Failed to sync plan with document. Please try again.")
    }
  }

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
      //const itemToDelete = planningItems.find((item) => item.id === id)
      //if (itemToDelete && itemToDelete.isDefault) {
      //  console.log("Default sections cannot be deleted.")
      //  return
      //}

      setPlanningItems((prev) => prev.filter((item) => item.id !== id))
      console.log("968-planning items:", planningItems.length)
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
    setBuildingToc(true)

    try {
      // Filter out items that are not selected (for default items)
      const itemsToAdd = tocItems.filter((item) => !item.isDefault || selectedTemplateItems[item.id] !== false)

      // Find items that are in TOC but not in plan
      const itemsToAddToPlan = itemsToAdd.filter(
        (tocItem) => !planningItems.some((planItem) => planItem.id === tocItem.id),
      )

      if (itemsToAddToPlan.length === 0) {
        setError("All selected TOC items are already in the plan.")
        setBuildingToc(false)
        return
      }

      // Create planning items from TOC items
      const newPlanningItems = [...planningItems]

      for (const tocItem of itemsToAddToPlan) {
        const newItem = {
          ...tocItem,
          status: "empty",
          comments: getDefaultComment(tocItem.title),
          words: 0,
          paragraphs: 0,
          tables: 0,
          graphics: 0,
        }

        newPlanningItems.push(newItem)
      }

      setPlanningItems(newPlanningItems)

      // Save after update
      await saveToDocumentProperties()

      setError(`Added ${itemsToAddToPlan.length} items to the plan.`)

      // Switch to plan tab
      setActiveTab("plan")
    } catch (error) {
      console.error("Error creating plan from TOC:", error)
      setError("Failed to create plan from TOC. Please try again.")
    } finally {
      setBuildingToc(false)
    }
  }

  // Build document structure with headers
  const buildDocumentStructure = async () => {
    setBuildingDocument(true)

    try {
      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
        console.log("This feature requires the Word API, which is not available in this environment.")
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

          // Inside the buildDocumentStructure function, after the for loop that inserts items:
          // Update status of empty items to created
          const updatedItems = planningItems.map((item) =>
            item.status === "empty" ? { ...item, status: "created" } : item,
          )
          setPlanningItems(updatedItems)

          // Save the updated status
          setTimeout(() => saveToDocumentProperties(), 100)

          await context.sync()
          console.log("Document structure has been built with headers!")
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)
          console.log("Failed to build document structure. Please try again.")
        }
      })
    } catch (error) {
      console.error("Error building document structure:", error)
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
              //title={item.isDefault ? "Default sections cannot be edited" : "Click to edit"}
              title={"Click to edit"}
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

          {/* Replace the delete button in renderPlanningItem with: */}
          <TooltipHost content={"Long click to delete"}>
            <IconButton
              iconProps={{ iconName: "Delete" }}
              onMouseDown={() => handleDeleteMouseDown(item.id)}
              onMouseUp={handleDeleteMouseUp}
              onMouseLeave={handleDeleteMouseUp}
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

  // Add these functions inside the component, before the return statement (around line 1000)
  // Update TOC title
  const updateTocTitle = (id, title) => {
    try {
      setTocItems((prev) => prev.map((item) => (item.id === id ? { ...item, title } : item)))

      // Also update in planning items if it exists there
      setPlanningItems((prev) => prev.map((item) => (item.id === id ? { ...item, title } : item)))

      setEditingTocItem(null)

      // Save after update
      setTimeout(() => saveToDocumentProperties(), 100)
    } catch (error) {
      console.error("Error updating TOC title:", error)
      setError("Failed to update TOC title. Please try again.")
    }
  }

  // Delete TOC item
  const deleteTocItem = (id) => {
    try {
      // Check if the item exists in planning items
      const existsInPlan = planningItems.some((item) => item.id === id)

      // If it's a default item and exists in plan, don't delete
      if (tocItems.find((item) => item.id === id)?.isDefault && existsInPlan) {
        setError("Cannot delete default items that are in the plan.")
        return
      }

      setTocItems((prev) => prev.filter((item) => item.id !== id))

      // If it exists in planning items, also remove it from there
      if (existsInPlan) {
        setPlanningItems((prev) => prev.filter((item) => item.id !== id))
      }

      // Save after update
      setTimeout(() => saveToDocumentProperties(), 100)
    } catch (error) {
      console.error("Error deleting TOC item:", error)
      setError("Failed to delete TOC item. Please try again.")
    }
  }

  // Export TOC
  const exportToc = () => {
    try {
      // Create a JSON string from the TOC items
      const tocData = JSON.stringify(tocItems, null, 2)

      // Create a blob from the JSON string
      const blob = new Blob([tocData], { type: "application/json" })

      // Create a URL for the blob
      const url = URL.createObjectURL(blob)

      // Create a link element
      const link = document.createElement("a")
      link.href = url
      link.download = "toc-template.json"

      // Append the link to the body
      document.body.appendChild(link)

      // Click the link to trigger the download
      link.click()

      // Remove the link from the body
      document.body.removeChild(link)

      // Revoke the URL
      URL.revokeObjectURL(url)

      setError("TOC template exported successfully.")
    } catch (error) {
      console.error("Error exporting TOC:", error)
      setError("Failed to export TOC. Please try again.")
    }
  }

  // Import TOC
  const importToc = () => {
    try {
      // Create a file input element
      const input = document.createElement("input")
      input.type = "file"
      input.accept = ".json"

      // Add an event listener for when a file is selected
      input.onchange = (e) => {
        const file = e.target.files[0]
        if (!file) return

        const reader = new FileReader()
        reader.onload = (event) => {
          try {
            const tocData = JSON.parse(event.target.result)

            // Validate the data
            if (!Array.isArray(tocData)) {
              throw new Error("Invalid TOC data format.")
            }

            // Find the highest ID
            const highestId = Math.max(...tocData.map((item) => item.id || 0), 0)

            // Update the TOC items
            setTocItems(tocData)

            // Update the next ID
            setNextId(highestId + 1)

            // Initialize selected items
            const initialSelectedItems = {}
            tocData.forEach((item) => {
              initialSelectedItems[item.id] = true
            })
            setSelectedTemplateItems(initialSelectedItems)

            // Save after update
            setTimeout(() => saveToDocumentProperties(), 100)

            setError("TOC template imported successfully.")
          } catch (parseError) {
            console.error("Error parsing TOC data:", parseError)
            setError("Failed to parse TOC data. Please check the file format.")
          }
        }

        reader.readAsText(file)
      }

      // Click the input to open the file dialog
      input.click()
    } catch (error) {
      console.error("Error importing TOC:", error)
      setError("Failed to import TOC. Please try again.")
    }
  }

  // Save plan to file
  const savePlanToFile = () => {
    try {
      // Create a JSON string from the planning items
      const planData = JSON.stringify(
        {
          planningItems: planningItems,
          tocItems: tocItems,
        },
        null,
        2,
      )

      // Create a blob from the JSON string
      const blob = new Blob([planData], { type: "application/json" })

      // Create a URL for the blob
      const url = URL.createObjectURL(blob)

      // Create a link element
      const link = document.createElement("a")
      link.href = url
      link.download = "writing-plan.json"

      // Append the link to the body
      document.body.appendChild(link)

      // Click the link to trigger the download
      link.click()

      // Remove the link from the body
      document.body.removeChild(link)

      // Revoke the URL
      URL.revokeObjectURL(url)

      setError("Plan saved to file successfully.")
    } catch (error) {
      console.error("Error saving plan to file:", error)
      setError("Failed to save plan to file. Please try again.")
    }
  }

  // Load plan from file
  const loadPlanFromFile = () => {
    try {
      // Create a file input element
      const input = document.createElement("input")
      input.type = "file"
      input.accept = ".json"

      // Add an event listener for when a file is selected
      input.onchange = (e) => {
        const file = e.target.files[0]
        if (!file) return

        const reader = new FileReader()
        reader.onload = (event) => {
          try {
            const planData = JSON.parse(event.target.result)

            // Validate the data
            if (!planData.planningItems || !planData.tocItems) {
              throw new Error("Invalid plan data format.")
            }

            // Find the highest ID
            const highestId = Math.max(
              ...planData.planningItems.map((item) => item.id || 0),
              ...planData.tocItems.map((item) => item.id || 0),
              0,
            )

            // Update the planning and TOC items
            setPlanningItems(planData.planningItems)
            setTocItems(planData.tocItems)

            // Update the next ID
            setNextId(highestId + 1)

            // Hide the template button
            setShowTemplateButton(false)

            // Save after update
            setTimeout(() => saveToDocumentProperties(), 100)

            setError("Plan loaded from file successfully.")
          } catch (parseError) {
            console.error("Error parsing plan data:", parseError)
            setError("Failed to parse plan data. Please check the file format.")
          }
        }

        reader.readAsText(file)
      }

      // Click the input to open the file dialog
      input.click()
    } catch (error) {
      console.error("Error loading plan from file:", error)
      setError("Failed to load plan from file. Please try again.")
    }
  }

  // Confirm load plan from file
  const confirmLoadPlanFromFile = () => {
    if (window.confirm("Loading a plan from file will discard all existing plan data! Are you sure?")) {
      loadPlanFromFile()
    }
  }

  // Handle delete button mouse down
  const handleDeleteMouseDown = (id) => {
    // Set a timer for 250ms
    const timer = setTimeout(() => {
      deleteSection(id)
      setDeleteButtonTimer(null)
    }, 250)

    setDeleteButtonTimer(timer)
  }

  // Handle delete button mouse up
  const handleDeleteMouseUp = () => {
    // Clear the timer if mouse is released before timeout
    if (deleteButtonTimer) {
      clearTimeout(deleteButtonTimer)
      setDeleteButtonTimer(null)
    }
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

      {/* Replace the header Stack with: */}
      <Stack horizontal horizontalAlign="space-between" styles={headerStyles}>
        <Stack>
          <Stack horizontal verticalAlign="center">
            <h1 style={titleStyles.root}>Writing Planner</h1>
            <TooltipHost content="Help">
              <IconButton
                iconProps={{ iconName: "Help" }}
                onClick={() => window.open("https://www.writepro.app/help/wpb", "_blank")}
                styles={{ root: { height: 24, width: 24, marginLeft: 5 } }}
              />
            </TooltipHost>
          </Stack>
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
            {/* Inside the Planning tab, after the planning items list and before the Build Document button: */}
            {showTemplateButton && planningItems.length === 0 && (
              <PrimaryButton
                text="Start with Template"
                iconProps={{ iconName: "Template" }}
                onClick={createTemplateStructure}
                styles={{ root: { marginBottom: 10 } }}
              />
            )}

            <PrimaryButton
              text={buildingDocument ? "Building..." : "Build Document Structure"}
              iconProps={{ iconName: "BuildDefinition" }}
              onClick={buildDocumentStructure}
              disabled={buildingDocument}
            />

            <Stack horizontal tokens={{ childrenGap: 10, padding: "10px 0" }}>
              <DefaultButton
                text="Save Plan to File"
                iconProps={{ iconName: "Save" }}
                onClick={savePlanToFile}
                styles={{ root: { flexGrow: 1 } }}
              />
              <DefaultButton
                text="Load Plan from File"
                iconProps={{ iconName: "OpenFile" }}
                onClick={confirmLoadPlanFromFile}
                styles={{ root: { flexGrow: 1 } }}
              />
            </Stack>
          </Stack>
        </PivotItem>
        <PivotItem headerText="TOC Template" itemKey="toc" itemIcon="BulletedList">
          <Stack tokens={{ childrenGap: 10 }}>
            {/* Update the TOC tab rendering section (around line 1200-1240) // Replace the existing TOC items rendering
            with this: */}
            <div style={{ maxHeight: 200, overflowY: "auto" }}>
              {tocItems.map((item) => (
                <div
                  key={item.id}
                  style={{
                    paddingLeft: (item.level - 1) * 15,
                    marginBottom: 5,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center" }}>
                    {item.isDefault && (
                      <input
                        type="checkbox"
                        checked={selectedTemplateItems[item.id] !== false}
                        onChange={(e) => {
                          setSelectedTemplateItems((prev) => ({
                            ...prev,
                            [item.id]: e.target.checked,
                          }))
                        }}
                        style={{ marginRight: 5 }}
                      />
                    )}
                    {editingTocItem === item.id ? (
                      <TextField
                        defaultValue={item.title}
                        autoFocus
                        onBlur={(e) => {
                          if (e && e.target) {
                            updateTocTitle(item.id, e.target.value)
                          }
                        }}
                        onKeyDown={(e) => {
                          if (e && e.key === "Enter" && e.target) {
                            updateTocTitle(item.id, e.target.value)
                          }
                        }}
                        styles={{ root: { width: 150, minWidth: 80 } }}
                      />
                    ) : (
                      <span>{item.title}</span>
                    )}
                    {planningItems.find((p) => p.id === item.id)?.status === "empty" && (
                      <span style={{ marginLeft: 5, color: "#c19c00" }}>⚠</span>
                    )}
                    {item.isDefault && (
                      <span style={{ marginLeft: 5, fontSize: "10px", color: "#666" }}>(default)</span>
                    )}
                  </div>
                  <div>
                    <IconButton
                      iconProps={{ iconName: "Edit" }}
                      onClick={() => setEditingTocItem(item.id)}
                      styles={{ root: { height: 24, width: 24 } }}
                    />
                    <IconButton
                      iconProps={{ iconName: "Delete" }}
                      onClick={() => deleteTocItem(item.id)}
                      disabled={item.isDefault && planningItems.some((p) => p.id === item.id)}
                      styles={{ root: { height: 24, width: 24 } }}
                    />
                  </div>
                </div>
              ))}
            </div>
            {/* Update the buttons in the TOC tab (around line 1240) // Replace the existing buttons with these: */}
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
            <Stack horizontal tokens={{ childrenGap: 10, padding: "10px 0" }}>
              <PrimaryButton
                text={buildingToc ? "Creating..." : "Create Plan with TOC"}
                iconProps={{ iconName: "FileTemplate" }}
                onClick={createTocScaffold}
                disabled={buildingToc}
                styles={{ root: { flexGrow: 2 } }}
              />
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <DefaultButton
                text="Export TOC"
                iconProps={{ iconName: "Download" }}
                onClick={exportToc}
                styles={{ root: { flexGrow: 1 } }}
              />
              <DefaultButton
                text="Import TOC"
                iconProps={{ iconName: "Upload" }}
                onClick={importToc}
                styles={{ root: { flexGrow: 1 } }}
              />
            </Stack>
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
          subText: "Created By Ali Vakilzadeh (CC)2025 using V0.dev",
        }}
      >
        <div style={{ margin: "20px 0" }}>
          <p>Contact: ali.vakilzadeh@gmail.com</p>
          <p style={{ marginTop: 10 }}>
            Plan and structure your documents before writing. Monitor and control your progress before publishing.
          </p>
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
  )
}
