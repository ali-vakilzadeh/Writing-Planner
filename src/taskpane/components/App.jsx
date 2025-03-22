"use client"
import { useState, useEffect } from "react"
import { DefaultButton, PrimaryButton, IconButton } from "@fluentui/react/lib/Button"
import { Stack } from "@fluentui/react/lib/Stack"
import { TextField } from "@fluentui/react/lib/TextField"
import { Dropdown } from "@fluentui/react/lib/Dropdown"
import { Pivot, PivotItem } from "@fluentui/react/lib/Pivot"
import { ProgressIndicator } from "@fluentui/react/lib/ProgressIndicator"
import { DetailsList, DetailsListLayoutMode, SelectionMode } from "@fluentui/react/lib/DetailsList"
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
    maxWidth: 300,
    margin: "0 auto",
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

  const [tocItems, setTocItems] = useState([])
  const [planningItems, setPlanningItems] = useState([])
  const [activeTab, setActiveTab] = useState("plan")
  const [refreshing, setRefreshing] = useState(false)
  const [nextId, setNextId] = useState(1)
  const [dataLoaded, setDataLoaded] = useState(false)
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

  // Load data from document properties
  useEffect(() => {
    if (isOfficeInitialized && !dataLoaded) {
      loadFromDocumentProperties()
    }
  }, [isOfficeInitialized, dataLoaded])

  // Initialize with template structure if nothing was loaded
  useEffect(() => {
    if (dataLoaded && tocItems.length === 0 && planningItems.length === 0) {
      createTemplateStructure()
    }
  }, [dataLoaded, tocItems.length, planningItems.length])

  const loadFromDocumentProperties = async () => {
    try {
      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
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
          } else {
            // If no saved data, initialize with empty arrays
            setTocItems([])
            setPlanningItems([])
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
      setTocItems([])
      setPlanningItems([])
      setDataLoaded(true)
      setError("Failed to load data. Please try again.")
    }
  }

  const createTemplateStructure = () => {
    try {
      // Comprehensive document template structure
      const templateTocItems = [
        { id: 1, title: "Title Page", level: 1 },
        { id: 2, title: "Abstract", level: 1 },
        { id: 3, title: "Table of Contents", level: 1 },
        { id: 4, title: "List of Figures", level: 1 },
        { id: 5, title: "List of Tables", level: 1 },
        { id: 6, title: "Introduction", level: 1 },
        { id: 7, title: "Background", level: 2 },
        { id: 8, title: "Problem Statement", level: 2 },
        { id: 9, title: "Research Questions", level: 2 },
        { id: 10, title: "Significance of Study", level: 2 },
        { id: 11, title: "Literature Review", level: 1 },
        { id: 12, title: "Theoretical Framework", level: 2 },
        { id: 13, title: "Previous Research", level: 2 },
        { id: 14, title: "Research Gap", level: 2 },
        { id: 15, title: "Methodology", level: 1 },
        { id: 16, title: "Research Design", level: 2 },
        { id: 17, title: "Data Collection", level: 2 },
        { id: 18, title: "Data Analysis", level: 2 },
        { id: 19, title: "Ethical Considerations", level: 2 },
        { id: 20, title: "Results", level: 1 },
        { id: 21, title: "Primary Findings", level: 2 },
        { id: 22, title: "Secondary Findings", level: 2 },
        { id: 23, title: "Discussion", level: 1 },
        { id: 24, title: "Interpretation of Results", level: 2 },
        { id: 25, title: "Limitations", level: 2 },
        { id: 26, title: "Implications", level: 2 },
        { id: 27, title: "Conclusion", level: 1 },
        { id: 28, title: "Summary", level: 2 },
        { id: 29, title: "Future Research", level: 2 },
        { id: 30, title: "References", level: 1 },
        { id: 31, title: "Appendices", level: 1 },
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
      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
        // Save to localStorage for development
        localStorage.setItem(
          "documentPlannerData",
          JSON.stringify({
            tocItems,
            planningItems: planningItems.map((item) => ({
              id: item.id,
              title: item.title,
              level: item.level,
              status: item.status,
              comments: item.comments,
            })),
          }),
        )
        return
      }

      await Word.run(async (context) => {
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
            })),
          }

          // Get document properties
          const properties = context.document.properties.customProperties
          if (!properties) {
            console.error("Document properties not available")
            return
          }

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
        setDataLoaded(false) // This will trigger the loading process again
        setDeleteConfirmOpen(false)
        return
      }

      await Word.run(async (context) => {
        try {
          // Get document properties
          const properties = context.document.properties.customProperties
          if (!properties) {
            console.error("Document properties not available")
            return
          }

          // Delete our data property
          properties.delete("documentPlannerData")

          await context.sync()
          console.log("Data deleted successfully")

          // Reset the state
          setTocItems([])
          setPlanningItems([])
          setNextId(1)
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

  // In a real add-in, this would fetch actual statistics from the Word document
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
          // In a real implementation, you would get actual statistics from the document
          // For now, we'll simulate it

          const updatedItems = planningItems.map((item) => ({
            ...item,
            words: item.status !== "empty" ? Math.floor(Math.random() * 500) + 50 : 0,
            paragraphs: item.status !== "empty" ? Math.floor(Math.random() * 10) + 1 : 0,
            tables: Math.random() > 0.7 ? Math.floor(Math.random() * 3) : 0,
            graphics: Math.random() > 0.8 ? Math.floor(Math.random() * 2) : 0,
          }))

          setPlanningItems(updatedItems)
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)
        }
      })
    } catch (error) {
      console.error("Error refreshing statistics:", error)
      setError("Failed to refresh statistics. Please try again.")
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

      const completedSections = planningItems.filter((item) => item.status === "finalized").length

      return (completedSections / totalSections) * 100
    } catch (error) {
      console.error("Error calculating completion:", error)
      return 0
    }
  }

  // Verify plan against TOC
  const verifyPlan = () => {
    try {
      // In a real add-in, this would compare the actual TOC from Word
      // with the planning items to ensure all sections are accounted for
      const missingInPlan = tocItems.filter((tocItem) => !planningItems.some((planItem) => planItem.id === tocItem.id))

      const missingInToc = planningItems.filter((planItem) => !tocItems.some((tocItem) => tocItem.id === planItem.id))

      if (missingInPlan.length === 0 && missingInToc.length === 0) {
        alert("Plan and TOC are in sync!")
      } else {
        alert(
          `Found ${missingInPlan.length} items in TOC but not in plan, and ${missingInToc.length} items in plan but not in TOC.`,
        )
      }
    } catch (error) {
      console.error("Error verifying plan:", error)
      setError("Failed to verify plan. Please try again.")
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
      }

      setPlanningItems((prev) => [...prev, newItem])
      setTocItems((prev) => [...prev, { id: nextId, title: "New Section", level }])
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
    setBuildingToc(true)

    try {
      // Check if Word API is available
      if (!Word || typeof Word.run !== "function") {
        console.error("Word API is not available")
        alert("This feature requires the Word API, which is not available in this environment.")
        setBuildingToc(false)
        return
      }

      await Word.run(async (context) => {
        try {
          // Sort items by ID to maintain the correct order
          const sortedItems = [...tocItems].sort((a, b) => a.id - b.id)

          // Insert a title for the TOC
          const titleParagraph = context.document.body.insertParagraph("TABLE OF CONTENTS", "Start")
          if (titleParagraph && titleParagraph.font) {
            titleParagraph.font.set({
              bold: true,
              size: 16,
            })
          }

          // Insert each TOC item with appropriate indentation
          for (const item of sortedItems) {
            const indent = "  ".repeat(item.level - 1)
            const paragraph = context.document.body.insertParagraph(`${indent}${item.title}`, "End")

            // Set indentation based on level
            if (paragraph) {
              paragraph.leftIndent = (item.level - 1) * 20
            }
          }

          await context.sync()
          alert("TOC scaffold has been created in the document!")
        } catch (contextError) {
          console.error("Error in Word.run context:", contextError)
          alert("Failed to create TOC scaffold. Please try again.")
        }
      })
    } catch (error) {
      console.error("Error creating TOC scaffold:", error)
      alert("Failed to create TOC scaffold. Please try again.")
      setError("Failed to create TOC scaffold. Please try again.")
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

            // Insert a paragraph break after each header
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

  // Columns for the planning items list
  const planningColumns = [
    {
      key: "stats",
      name: "Stats",
      fieldName: "stats",
      minWidth: 30,
      maxWidth: 30,
      onRender: (item) => (
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
          />
        </TooltipHost>
      ),
    },
    {
      key: "title",
      name: "Section",
      fieldName: "title",
      minWidth: 100,
      onRender: (item) => (
        <div style={{ paddingLeft: (item.level - 1) * 15 }}>
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
            />
          ) : (
            <Stack horizontal tokens={{ childrenGap: 5 }}>
              <span>{item.title}</span>
              <IconButton iconProps={{ iconName: "Edit" }} onClick={() => setEditingItem(item.id)} />
              <IconButton iconProps={{ iconName: "Comment" }} onClick={() => setCommentItem(item.id)} />
            </Stack>
          )}
        </div>
      ),
    },
    {
      key: "status",
      name: "Status",
      fieldName: "status",
      minWidth: 80,
      maxWidth: 80,
      onRender: (item) => (
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
              width: 80,
            },
            title: {
              backgroundColor: STATUS_COLORS[item.status]?.background,
              color: STATUS_COLORS[item.status]?.color,
              borderColor: "transparent",
            },
          }}
        />
      ),
    },
    {
      key: "delete",
      name: "",
      fieldName: "delete",
      minWidth: 30,
      maxWidth: 30,
      onRender: (item) => (
        <TooltipHost content="Delete Section">
          <IconButton iconProps={{ iconName: "Delete" }} onClick={() => deleteSection(item.id)} />
        </TooltipHost>
      ),
    },
  ]

  // If Office is not initialized
  if (!isOfficeInitialized) {
    return (
      <Stack styles={containerStyles}>
        <Spinner label="Loading Office.js..." size={SpinnerSize.large} />
      </Stack>
    )
  }

  return (
    <Stack styles={containerStyles}>
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
        <DefaultButton text="Verify" iconProps={{ iconName: "CheckMark" }} onClick={verifyPlan} />
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
            <DetailsList
              items={planningItems}
              columns={planningColumns}
              layoutMode={DetailsListLayoutMode.fixedColumns}
              selectionMode={SelectionMode.none}
              compact={true}
            />
            <PrimaryButton
              text={buildingDocument ? "Building..." : "Build Document Structure"}
              iconProps={{ iconName: "BuildDefinition" }}
              onClick={buildDocumentStructure}
              disabled={buildingDocument}
            />
          </Stack>
        </PivotItem>
        <PivotItem headerText="TOC" itemKey="toc" itemIcon="BulletedList">
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
                </div>
              ))}
            </div>
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <DefaultButton text="Add L1" iconProps={{ iconName: "Add" }} onClick={() => addSection(1)} />
              <DefaultButton text="Add L2" iconProps={{ iconName: "Add" }} onClick={() => addSection(2)} />
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
  )
}