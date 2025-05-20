# NET.ALR - Network Analysis and Monitoring Tool

![NET.ALR Logo](https://via.placeholder.com/150x150)

## Overview

NET.ALR is an advanced network monitoring and analysis tool developed in Excel VBA. This application provides comprehensive network device tracking, IP classification, statistical analysis, and data visualization capabilities through an intuitive user interface.

## Features

### Authentication System
- Secure login interface
- User credential verification

### Multi-Page Dashboard Interface
- Tab-based navigation system
- Clean and professional UI design

### Advanced Search Functionality
- Multi-criteria search options (IP, Device Name, Type, Status)
- Real-time result filtering
- Dynamic detail display

### IP Classification System
- Automatic IP class detection (A, B, C, D, E)
- Organized ListBox display by class
- IP validation and hostname association

### Statistical Analysis Tools
- Consumption analysis by device type
- Operating system distribution tracking
- Priority distribution monitoring
- Active vs. inactive device tracking

### Data Visualization
- Consumption vs. Latency charts
- Device type distribution pie charts
- Network priority bar graphs
- Auto-updating graphs

### Network Data Management
- Device tracking (IP, Name, Type, Status)
- Performance monitoring (Consumption, Latency)
- Security assessment (Open ports)
- System information tracking (OS, Update status)

## Screenshots

![Login Interface](https://via.placeholder.com/500x300)
*Login Interface*

![Main Dashboard](https://via.placeholder.com/500x300)
*Main Application Dashboard*

## Installation & Setup

1. **Requirements**
   - Microsoft Excel (2016 or later recommended)
   - Macros must be enabled

2. **Installation Steps**
   - Download the NET.ALR.xlsm file
   - Open in Excel
   - Enable macros when prompted
   - Login with default credentials (Username: Professeur, Password: Laatabi)

## Usage Guide

### Login
- Enter username and password in the login form
- Click "Login" to access the main dashboard

### Device Search
1. Navigate to the Search tab
2. Enter search criteria (IP, Name, Type, or Status)
3. Review matching results in the list
4. Select a device to view detailed information

### IP Classification
1. Navigate to the IP Classification tab
2. View automatically sorted IPs by class
3. Select an IP to view associated device details

### Statistical Analysis
1. Navigate to the Analysis tab
2. View consumption statistics by device type
3. Review OS distribution data
4. Analyze network priority allocation

### Data Visualization
1. Navigate to the Visualization tab
2. Interact with the various charts for different insights
3. Use filters to customize the view if available

## Project Structure

- `NET.ALR.xlsm` - Main application file
- `Modules/` - VBA code modules
  - `modMain.bas` - Core functionality
  - `modDataHandling.bas` - Data processing functions
  - `modGraphics.bas` - Chart generation code
- `Forms/` - VBA userforms
  - `frmLogin.frm` - Login interface
  - `frmMain.frm` - Main application interface

## Technical Implementation

### Technologies Used
- Excel VBA
- Excel Tables (ListObjects)
- Excel Charts
- ActiveX Controls
- UserForms
- MultiPage Controls

### Advanced Techniques
- Error handling for robust operation
- Dictionary objects for efficient data storage
- String manipulation for IP processing
- Dynamic chart generation and conversion

## Contributors

- Badr Laklach
- Imane Ait Boukdir
- Alae Rhazouani

## Acknowledgements

Special thanks to Monsieur Laatabi for guidance and supervision throughout the development of this project.

## License

This project is released under the MIT License. See the LICENSE file for details.

---

© 2024 NET.ALR Team | Made for GESTION ET ANALYSE DES DONNEES course | RESEAUX INFORMATIQUES
