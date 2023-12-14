#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#


library(tidyverse)
library(ggplot2)
library(readxl)
library(shiny)
library(shinydashboard)
library(lubridate)
library(reactable)


# Set the directory path where your Excel files are located

directory_path <- "./data/"

# Get a list of all xlsx files in the directory
xlsx_files <- list.files(path = directory_path, pattern = "\\.xlsx$", full.names = TRUE)

# Specify the sheet name you want to read
sheet_name <- "Products"

# Read all Excel files and combine them into a single DataFrame
combined_data <- map_dfr(xlsx_files, ~read_excel(.x, , sheet = sheet_name, skip = 2)) %>%
  mutate_all(~replace(., is.nan(.), 0)) %>% #View()
  mutate(`Pret mediu de vanzare (RON)` = as.numeric(`Pret mediu de vanzare (RON)`),
         `Cantitate vanduta` = ifelse(`Cantitate vanduta` == '-', 0, `Cantitate vanduta`),
         `Cantitate vanduta` = as.numeric(`Cantitate vanduta`)) 
 

#combined_data %>% View()

#########################################################################################

# Define UI for application that draws a histogram
sidebar <- dashboardSidebar(
  #width = 200,
  fluidPage(
    titlePanel(h4(strong("Filters"))),
    column(5,checkboxGroupInput("An", label = "An",
                         sort(unique(combined_data$An)), selected = format(Sys.Date(), "%Y"),
                         )),
    
    column(5,checkboxGroupInput("Luna", label = "Luna",
                         choices = sort(unique(combined_data$Luna)), selected = max(unique(combined_data$Luna)))),
    
    
  ),
  fluidPage(
    titlePanel(h4(strong("KPIs"))),
  ),
  sidebarMenu(
    menuItem("  Visits", tabName = "visits", icon = icon("globe")),
    menuItem("  Sales",  tabName = "sales", icon = icon("dollar-sign"))
  )
)

body <- dashboardBody(
  tabItems(
    tabItem(tabName = "visits",
            h2("Visits per Month"),
            tabsetPanel(
              type = "tabs",
              tabPanel("Plot", plotOutput("visitplot")),
              tabPanel("Data", uiOutput("visittable"))
              )
    ),
    
    tabItem(tabName = "sales",
            h2("Sales per Month"),
            tabsetPanel(
              type = "tabs",
              tabPanel("Plot", plotOutput("vanzariplot")),
              tabPanel("Data", uiOutput("vanzaritable"))
            )
    )
  )
)

# Put them together into a dashboardPage
ui <- dashboardPage(
  skin = "green",
  dashboardHeader(title = h3("SELAMAX KPIs")),
  sidebar,
  body
)



######################################################################
# Define server logic required to draw a histogram
server <- function(input, output) {
 
   filtered_data <- reactive({
    combined_data %>%
      filter(( An %in% input$An),
             ( Luna %in% input$Luna)) 
    
  })
  
  
    output$visitplot <- renderPlot({
      filtered_data() %>% 
        #combined_data%>%
        ggplot(aes(x = interaction(An, Luna), y = Vizite, group = `Nume produs`, color = `Nume produs`)) +
        geom_line(show.legend = FALSE) +
        geom_point(show.legend = FALSE) +
        facet_wrap(~str_wrap(`Nume produs`, width = 70)) +
        labs(title = "Number of Visits per Month for Each Product",
             x = "Month-Year",
             y = "Number of Visits") +
        #scale_fill_discrete(guide="none")+
        theme_bw() + 
        theme(
          plot.title = element_text(face = "bold", size = 20),
          strip.text = element_text(size = 12, angle = 0),  # Adjust hjust as needed
          strip.background = element_blank(),
          axis.title.y = element_text(face="bold"),
          axis.title.x = element_text( face="bold"),
          axis.text.x = element_text(size= 10,face="bold")
        ) +
        # Adjust the size of the plotting area to accommodate longer titles
        coord_cartesian(clip = 'off')
      
    },height=800
    )
    
    
    output$visittable <- renderUI({
      
      columns_to_display <- c("An", "Luna","Perioada", "Nume produs", "Vizite")
      
      reactable(
        filtered_data()[, columns_to_display, drop = FALSE] ,
        filterable = TRUE,
        searchable = TRUE,
        showPageSizeOptions = TRUE,
        defaultSorted = list("Nume produs" = "asc", "An" = "desc", "Luna" = "desc"),
        groupBy = c("Nume produs", "An"),
        columns = list(
          An = colDef(  width = 70),
          Luna = colDef( width = 70),
          Perioada = colDef(aggregate = "count",format = list(aggregated = colFormat(suffix = " months")), width = 150),
          Vizite = colDef(aggregate = "mean", format = colFormat(digits = 0), width = 100),
          "Nume produs" = colDef(rowHeader = TRUE, style = list(fontWeight = 600)))
      
      )
    })
    
    
    output$vanzariplot <- renderPlot({
      
      max_cantitate <- max(filtered_data()$`Cantitate vanduta`, na.rm = TRUE)*1.1
      max_pret_mediu <-max(combined_data$`Pret mediu de vanzare (RON)`, na.rm = TRUE)*1.1
      coef <- max_pret_mediu/max_cantitate
      
      filtered_data() %>% 
        #combined_data%>%
        ggplot(aes(x = interaction(An, Luna),  group = `Nume produs`, color = `Nume produs`)) +
        geom_bar(aes(y = `Cantitate vanduta`, fill = `Nume produs`),stat="identity", alpha = 0.7, show.legend = FALSE) +
        geom_line(aes(y = `Pret mediu de vanzare (RON)`/coef), color = "red") +  
        geom_point(aes(y = `Pret mediu de vanzare (RON)`/coef), color = "red") +  
        scale_y_continuous(name = "Cantitate vanduta",  limits = c(0, max_cantitate),breaks = seq(0, max_cantitate, by=2),
                           sec.axis = sec_axis(~.*coef, name = "Pret mediu de vanzare (RON)")) +
        facet_wrap(~str_wrap(`Nume produs`, width = 70)) +
        labs(title = "Number of sold product and average price per Month",
             x = "Month",
             y = "Number of Visits") +
        #scale_fill_discrete(guide="none")+
        theme_bw() +
        theme(
          plot.title = element_text(face = "bold", size = 20),
          strip.text = element_text(size = 12, angle = 0),  # Adjust hjust as needed
          strip.background = element_blank(),
          axis.title.y.right = element_text(color = "red", face="bold"),
          axis.text.y.right=element_text(color="red"),
          axis.title.y.left = element_text(face="bold"),
          axis.title.x = element_text(face="bold"),
          axis.text.x = element_text(size= 10, face="bold")
        ) +
        # Adjust the size of the plotting area to accommodate longer titles
        coord_cartesian(clip = 'off')
    },height=800
    )
    
    
    output$vanzaritable <- renderUI({
      
      columns_to_display <- c("An", "Luna","Perioada", "Nume produs", "Cantitate vanduta", "Pret mediu de vanzare (RON)")
      
      reactable(
        filtered_data()[, columns_to_display, drop = FALSE] ,
        filterable = TRUE,
        searchable = TRUE,
        showPageSizeOptions = TRUE,
        defaultSorted = list("Nume produs" = "asc", "An" = "desc", "Luna" = "desc"),
        groupBy = c("Nume produs", "An"),
        columns = list(
          An = colDef(  width = 70),
          Luna = colDef( width = 70),
          Perioada = colDef(aggregate = "count",format = list(aggregated = colFormat(suffix = " months")), width = 150),
          "Cantitate vanduta" = colDef(aggregate = "mean", format = colFormat(digits = 0), width = 100),
          "Pret mediu de vanzare (RON)" = colDef(aggregate = "mean", format = colFormat(digits = 0), width = 100),
          "Nume produs" = colDef(rowHeader = TRUE, style = list(fontWeight = 600)))
        
      )
      
      
    
    })
}

# Run the application 
shinyApp(ui = ui, server = server)
