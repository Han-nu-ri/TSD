library("shiny")

# Define UI
ui <- fluidPage(
  titlePanel("PCA"),

  sidebarLayout(
    # SideBar Component Info
    # 1. fileInput
    # 2. X Label Select UI
    # 3. Y Label Select UI
    sidebarPanel(
      fileInput("file", label = h3("File Input")),
      htmlOutput("selectXLabel"),
      htmlOutput("selectYLabel")
    ),
    
    # 1. plot graph
    mainPanel(
      plotOutput("pcaPlot")
    )
  )
  
)

# Define Server Logic
server <- function(input, output) {
  # Make Select X Label UI Dynamically
  output$selectXLabel <- renderUI({
    selectInput("xLabel", "Select X Label", list("Percent White", 
                                                 "Percent Black",
                                                 "Percent Hispanic", 
                                                 "Percent Asian"))
    
  })
  
  # Make Select Y Label UI Dynamically
  output$selectYLabel <- renderUI({
    selectInput("yLabel", "Select Y Label", list("Percent White", 
                                                 "Percent Black",
                                                 "Percent Hispanic", 
                                                 "Percent Asian"))
    
  })
  
  # Draw Plot
  output$pcaPlot <- renderPlot(
    plot(x = c(1:10), y = c(2:11))
  )
}

# Run the app
shinyApp(ui = ui, server = server)