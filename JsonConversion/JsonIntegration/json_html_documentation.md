# JSON to HTML Integration Documentation

## Overview

This document outlines various approaches for loading and displaying JSON data in HTML pages, with specific focus on local file scenarios and Microsoft Access integration using the Edge browser control. The solution evolved from exploring external JSON file loading to implementing embedded JavaScript data structures.

## User Requirements and Prompts

### Initial Requirements
1. **Edge Browser Control**: How to demonstrate an HTML page located on the C drive in the public folder using the Edge browser
2. **VBA Integration**: How to control the Edge browser from VBA in MS Access using the edge browser control
3. **JSON Loading**: How to load a JSON file into an HTML page located in the same directory as the HTML page
4. **Embedded Approach**: Request to embed the JSON as JavaScript instead of external file loading
5. **Example Data**: Request for a comprehensive example data.js file

### Evolution of Requirements
- Started with basic HTML file display in Edge
- Progressed to VBA/Access integration
- Moved to JSON data integration
- Settled on embedded JavaScript approach for reliability

## Approaches for Loading JSON Data

### Method 1: Using Fetch API (Modern Approach)
**Best for**: Modern browsers with network access
**Limitations**: CORS issues with local files

```javascript
fetch('./data.json')
    .then(response => response.json())
    .then(data => {
        // Use the data
        console.log(data);
    })
    .catch(error => {
        console.error('Error loading JSON:', error);
    });
```

### Method 2: Using XMLHttpRequest (Legacy Compatible)
**Best for**: Older browser compatibility
**Limitations**: More verbose, still has CORS issues

```javascript
function loadJSON() {
    const xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function() {
        if (xhr.readyState === 4 && xhr.status === 200) {
            const data = JSON.parse(xhr.responseText);
            displayData(data);
        }
    };
    xhr.open('GET', './data.json', true);
    xhr.send();
}
```

### Method 3: Using jQuery
**Best for**: Projects already using jQuery
**Limitations**: Requires jQuery library, CORS issues

```javascript
$.getJSON('./data.json', function(data) {
    console.log(data);
}).fail(function() {
    console.log('Error loading JSON file');
});
```

### Method 4: Embedded JavaScript (Recommended Solution)
**Best for**: Local files, Access integration, reliability
**Advantages**: No CORS issues, immediate availability, works offline

```javascript
// data.js file
const jsonData = {
    "key": "value",
    "array": [1, 2, 3]
};
```

```html
<!-- HTML file -->
<script src="./data.js"></script>
<script>
    // Data is immediately available
    console.log(jsonData);
</script>
```

## VBA Integration with MS Access

### WebView2 Control (Recommended)
```vb
Private Sub Form_Load()
    ' Navigate to HTML file in Public folder
    Me.WebView21.Navigate "file:///C:/Users/Public/yourfile.html"
End Sub

Private Sub btnNavigate_Click()
    Dim filePath As String
    filePath = "file:///C:/Users/Public/demo.html"
    Me.WebView21.Navigate filePath
End Sub
```

### Legacy WebBrowser Control
```vb
Private Sub Form_Load()
    Me.WebBrowser0.Navigate "file:///C:/Users/Public/yourfile.html"
End Sub
```

### External Edge Launch
```vb
Private Sub btnOpenInEdge_Click()
    Dim edgePath As String
    Dim htmlFile As String
    
    edgePath = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"""
    htmlFile = """file:///C:/Users/Public/yourfile.html"""
    
    Shell edgePath & " " & htmlFile, vbNormalFocus
End Sub
```

## Complete Implementation Example

### File Structure
```
C:\Users\Public\
├── index.html
└── data.js
```

### HTML Implementation (index.html)
```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Company Dashboard</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #007acc;
            color: white;
        }
        .button {
            background-color: #007acc;
            color: white;
            padding: 10px 15px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1 id="companyName">Loading...</h1>
        <button class="button" onclick="showAllEmployees()">Show Employees</button>
        <button class="button" onclick="showDepartments()">Show Departments</button>
        <button class="button" onclick="showProjects()">Show Projects</button>
        <div id="content"></div>
    </div>

    <!-- Include the data.js file -->
    <script src="./data.js"></script>
    
    <script>
        window.onload = function() {
            document.getElementById('companyName').textContent = jsonData.company.name;
        };

        function showAllEmployees() {
            let html = '<h3>All Employees</h3>';
            html += '<table><tr><th>Name</th><th>Position</th><th>Department</th><th>Salary</th></tr>';
            
            jsonData.employees.forEach(emp => {
                html += `<tr>
                    <td>${emp.fullName}</td>
                    <td>${emp.position}</td>
                    <td>${emp.department}</td>
                    <td>$${emp.salary.toLocaleString()}</td>
                </tr>`;
            });
            
            html += '</table>';
            document.getElementById('content').innerHTML = html;
        }

        function showDepartments() {
            let html = '<h3>Departments</h3>';
            html += '<table><tr><th>Name</th><th>Head</th><th>Budget</th><th>Employees</th></tr>';
            
            jsonData.departments.forEach(dept => {
                const head = DataUtils.getEmployeeById(dept.head);
                const employees = DataUtils.getEmployeesByDepartment(dept.name);
                
                html += `<tr>
                    <td>${dept.name}</td>
                    <td>${head ? head.fullName : 'N/A'}</td>
                    <td>$${dept.budget.toLocaleString()}</td>
                    <td>${employees.length}</td>
                </tr>`;
            });
            
            html += '</table>';
            document.getElementById('content').innerHTML = html;
        }

        function showProjects() {
            let html = '<h3>Projects</h3>';
            html += '<table><tr><th>Name</th><th>Status</th><th>Budget</th><th>Progress</th></tr>';
            
            jsonData.projects.forEach(project => {
                const progress = DataUtils.getProjectProgress(project.id);
                
                html += `<tr>
                    <td>${project.name}</td>
                    <td>${project.status}</td>
                    <td>$${project.budget.toLocaleString()}</td>
                    <td>${progress}%</td>
                </tr>`;
            });
            
            html += '</table>';
            document.getElementById('content').innerHTML = html;
        }
    </script>
</body>
</html>
```

## Data Structure (data.js)

### Company Information
```javascript
const jsonData = {
    "company": {
        "name": "TechFlow Solutions",
        "founded": 2018,
        "headquarters": "West Chester, PA",
        "industry": "Software Development",
        "website": "www.techflowsolutions.com",
        "phone": "(610) 555-0123",
        "email": "info@techflowsolutions.com",
        "taxId": "12-3456789",
        "lastUpdated": "2025-08-09"
    },
    // ... additional data structures
};
```

### Employee Records
```javascript
"employees": [
    {
        "id": 101,
        "firstName": "Sarah",
        "lastName": "Johnson",
        "fullName": "Sarah Johnson",
        "position": "CEO",
        "department": "Executive",
        "email": "sarah.johnson@techflow.com",
        "phone": "(610) 555-0101",
        "hireDate": "2018-01-15",
        "salary": 150000,
        "isActive": true,
        "manager": null,
        "skills": ["Leadership", "Strategy", "Business Development"],
        "address": {
            "street": "123 Main St",
            "city": "West Chester",
            "state": "PA",
            "zipCode": "19380"
        }
    }
    // ... additional employee records
]
```

### Department Structure
```javascript
"departments": [
    {
        "id": 1,
        "name": "Executive",
        "description": "Executive leadership and strategic planning",
        "head": 101,
        "budget": 300000,
        "location": "Building A, Floor 3",
        "employeeCount": 1
    }
    // ... additional departments
]
```

### Project Management
```javascript
"projects": [
    {
        "id": 1001,
        "name": "Customer Portal Redesign",
        "description": "Complete redesign of the customer-facing portal",
        "status": "In Progress",
        "priority": "High",
        "startDate": "2025-06-01",
        "endDate": "2025-10-15",
        "budget": 150000,
        "spentBudget": 75000,
        "teamLead": 103,
        "assignedEmployees": [102, 103, 106],
        "milestones": [
            {
                "name": "Requirements Gathering",
                "dueDate": "2025-06-30",
                "completed": true
            }
            // ... additional milestones
        ]
    }
    // ... additional projects
]
```

## Utility Functions

### Employee Management
```javascript
const DataUtils = {
    getEmployeeById: function(id) {
        return jsonData.employees.find(emp => emp.id === id);
    },

    getEmployeesByDepartment: function(department) {
        return jsonData.employees.filter(emp => emp.department === department);
    },

    getActiveEmployees: function() {
        return jsonData.employees.filter(emp => emp.isActive);
    },

    getDepartmentSalaryTotal: function(department) {
        const employees = this.getEmployeesByDepartment(department);
        return employees.reduce((total, emp) => total + emp.salary, 0);
    }
};
```

### Project Management
```javascript
getProjectById: function(id) {
    return jsonData.projects.find(project => project.id === id);
},

getProjectsByStatus: function(status) {
    return jsonData.projects.filter(project => project.status === status);
},

getProjectProgress: function(projectId) {
    const project = this.getProjectById(projectId);
    if (!project) return 0;
    
    const completedMilestones = project.milestones.filter(m => m.completed).length;
    return Math.round((completedMilestones / project.milestones.length) * 100);
}
```

### Financial Operations
```javascript
formatCurrency: function(amount) {
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD'
    }).format(amount);
},

getYearToDateRevenue: function() {
    return jsonData.financials.quarterlyRevenue.reduce((total, quarter) => 
        total + quarter.revenue, 0);
}
```

## Usage Examples

### Basic Data Access
```javascript
// Get company name
const companyName = jsonData.company.name;

// Get all employees
const allEmployees = jsonData.employees;

// Find specific employee
const employee = DataUtils.getEmployeeById(101);
```

### Complex Operations
```javascript
// Calculate department costs
const techSalaries = DataUtils.getDepartmentSalaryTotal('Technology');

// Get project team
const team = DataUtils.getProjectTeam(1001);

// Search functionality
const results = DataUtils.searchEmployees('sarah');
```

### Dynamic HTML Generation
```javascript
function createEmployeeTable() {
    let html = '<table><tr><th>Name</th><th>Department</th></tr>';
    
    jsonData.employees.forEach(emp => {
        html += `<tr><td>${emp.fullName}</td><td>${emp.department}</td></tr>`;
    });
    
    html += '</table>';
    return html;
}
```

## Advantages of Embedded JavaScript Approach

1. **No CORS Issues**: Works with local files without browser security restrictions
2. **Immediate Availability**: Data loads instantly with the page
3. **Reliability**: No network requests that could fail
4. **Offline Support**: Works without internet connection
5. **Access Integration**: Perfect for MS Access WebView2 control
6. **Extensibility**: Easy to add utility functions and data processing
7. **Performance**: No HTTP requests or parsing delays

## Best Practices

1. **Data Organization**: Structure data logically with clear relationships
2. **Utility Functions**: Create helper functions for common operations
3. **Error Handling**: Include validation and error checking
4. **Documentation**: Comment complex data structures and functions
5. **Modularity**: Separate data, utilities, and presentation logic
6. **Consistency**: Use consistent naming conventions and data formats

## Integration with MS Access

### Setup Steps
1. Add WebView2 control to Access form
2. Place HTML and JS files in accessible directory (e.g., C:\Users\Public\)
3. Configure VBA to navigate to HTML file
4. Implement refresh and navigation methods

### Example VBA Implementation
```vb
Private Sub Form_Load()
    Me.WebView21.Navigate "file:///C:/Users/Public/index.html"
End Sub

Private Sub btnRefresh_Click()
    Me.WebView21.Reload()
End Sub

Private Sub NavigateToLocalHTML()
    On Error GoTo ErrorHandler
    
    Dim htmlPath As String
    htmlPath = "file:///C:/Users/Public/demo.html"
    
    If Dir("C:\Users\Public\demo.html") <> "" Then
        Me.WebView21.Navigate htmlPath
    Else
        MsgBox "HTML file not found!", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading HTML file: " & Err.Description, vbCritical
End Sub
```

## Complete data.js Example

```javascript
// data.js - Comprehensive business data example
const jsonData = {
    // Company Information
    "company": {
        "name": "TechFlow Solutions",
        "founded": 2018,
        "headquarters": "West Chester, PA",
        "industry": "Software Development",
        "website": "www.techflowsolutions.com",
        "phone": "(610) 555-0123",
        "email": "info@techflowsolutions.com",
        "taxId": "12-3456789",
        "lastUpdated": "2025-08-09"
    },

    // Employee Data
    "employees": [
        {
            "id": 101,
            "firstName": "Sarah",
            "lastName": "Johnson",
            "fullName": "Sarah Johnson",
            "position": "CEO",
            "department": "Executive",
            "email": "sarah.johnson@techflow.com",
            "phone": "(610) 555-0101",
            "hireDate": "2018-01-15",
            "salary": 150000,
            "isActive": true,
            "manager": null,
            "skills": ["Leadership", "Strategy", "Business Development"],
            "address": {
                "street": "123 Main St",
                "city": "West Chester",
                "state": "PA",
                "zipCode": "19380"
            }
        },
        {
            "id": 102,
            "firstName": "Michael",
            "lastName": "Chen",
            "fullName": "Michael Chen",
            "position": "CTO",
            "department": "Technology",
            "email": "michael.chen@techflow.com",
            "phone": "(610) 555-0102",
            "hireDate": "2018-02-01",
            "salary": 140000,
            "isActive": true,
            "manager": 101,
            "skills": ["JavaScript", "Python", "Cloud Architecture", "Team Leadership"],
            "address": {
                "street": "456 Oak Ave",
                "city": "Malvern",
                "state": "PA",
                "zipCode": "19355"
            }
        },
        {
            "id": 103,
            "firstName": "Emily",
            "lastName": "Rodriguez",
            "fullName": "Emily Rodriguez",
            "position": "Senior Developer",
            "department": "Technology",
            "email": "emily.rodriguez@techflow.com",
            "phone": "(610) 555-0103",
            "hireDate": "2019-03-15",
            "salary": 95000,
            "isActive": true,
            "manager": 102,
            "skills": ["React", "Node.js", "MongoDB", "AWS"],
            "address": {
                "street": "789 Pine St",
                "city": "Exton",
                "state": "PA",
                "zipCode": "19341"
            }
        },
        {
            "id": 104,
            "firstName": "David",
            "lastName": "Wilson",
            "fullName": "David Wilson",
            "position": "Product Manager",
            "department": "Product",
            "email": "david.wilson@techflow.com",
            "phone": "(610) 555-0104",
            "hireDate": "2020-06-01",
            "salary": 85000,
            "isActive": true,
            "manager": 101,
            "skills": ["Product Strategy", "Agile", "Analytics", "UX Design"],
            "address": {
                "street": "321 Elm Dr",
                "city": "Wayne",
                "state": "PA",
                "zipCode": "19087"
            }
        },
        {
            "id": 105,
            "firstName": "Jessica",
            "lastName": "Brown",
            "fullName": "Jessica Brown",
            "position": "Marketing Manager",
            "department": "Marketing",
            "email": "jessica.brown@techflow.com",
            "phone": "(610) 555-0105",
            "hireDate": "2021-01-10",
            "salary": 75000,
            "isActive": true,
            "manager": 101,
            "skills": ["Digital Marketing", "Content Strategy", "SEO", "Analytics"],
            "address": {
                "street": "654 Maple Ln",
                "city": "King of Prussia",
                "state": "PA",
                "zipCode": "19406"
            }
        },
        {
            "id": 106,
            "firstName": "Robert",
            "lastName": "Davis",
            "fullName": "Robert Davis",
            "position": "Junior Developer",
            "department": "Technology",
            "email": "robert.davis@techflow.com",
            "phone": "(610) 555-0106",
            "hireDate": "2023-08-15",
            "salary": 65000,
            "isActive": true,
            "manager": 103,
            "skills": ["JavaScript", "HTML/CSS", "Git", "SQL"],
            "address": {
                "street": "987 Cedar Way",
                "city": "Phoenixville",
                "state": "PA",
                "zipCode": "19460"
            }
        }
    ],

    // Department Information
    "departments": [
        {
            "id": 1,
            "name": "Executive",
            "description": "Executive leadership and strategic planning",
            "head": 101,
            "budget": 300000,
            "location": "Building A, Floor 3",
            "employeeCount": 1
        },
        {
            "id": 2,
            "name": "Technology",
            "description": "Software development and IT operations",
            "head": 102,
            "budget": 800000,
            "location": "Building A, Floor 2",
            "employeeCount": 3
        },
        {
            "id": 3,
            "name": "Product",
            "description": "Product management and strategy",
            "head": 104,
            "budget": 250000,
            "location": "Building A, Floor 1",
            "employeeCount": 1
        },
        {
            "id": 4,
            "name": "Marketing",
            "description": "Marketing and customer acquisition",
            "head": 105,
            "budget": 200000,
            "location": "Building B, Floor 1",
            "employeeCount": 1
        }
    ],

    // Project Data
    "projects": [
        {
            "id": 1001,
            "name": "Customer Portal Redesign",
            "description": "Complete redesign of the customer-facing portal with modern UI/UX",
            "status": "In Progress",
            "priority": "High",
            "startDate": "2025-06-01",
            "endDate": "2025-10-15",
            "budget": 150000,
            "spentBudget": 75000,
            "teamLead": 103,
            "assignedEmployees": [102, 103, 106],
            "milestones": [
                {
                    "name": "Requirements Gathering",
                    "dueDate": "2025-06-30",
                    "completed": true
                },
                {
                    "name": "UI/UX Design",
                    "dueDate": "2025-07-31",
                    "completed": true
                },
                {
                    "name": "Development Phase 1",
                    "dueDate": "2025-09-15",
                    "completed": false
                },
                {
                    "name": "Testing & Launch",
                    "dueDate": "2025-10-15",
                    "completed": false
                }
            ]
        },
        {
            "id": 1002,
            "name": "Mobile App Development",
            "description": "Native mobile app for iOS and Android platforms",
            "status": "Planning",
            "priority": "Medium",
            "startDate": "2025-09-01",
            "endDate": "2026-03-01",
            "budget": 200000,
            "spentBudget": 0,
            "teamLead": 102,
            "assignedEmployees": [102, 103, 104],
            "milestones": [
                {
                    "name": "Market Research",
                    "dueDate": "2025-09-30",
                    "completed": false
                },
                {
                    "name": "Technical Planning",
                    "dueDate": "2025-10-31",
                    "completed": false
                },
                {
                    "name": "MVP Development",
                    "dueDate": "2025-12-31",
                    "completed": false
                },
                {
                    "name": "Full Launch",
                    "dueDate": "2026-03-01",
                    "completed": false
                }
            ]
        },
        {
            "id": 1003,
            "name": "Marketing Automation",
            "description": "Implement marketing automation tools and workflows",
            "status": "Completed",
            "priority": "Medium",
            "startDate": "2025-03-01",
            "endDate": "2025-07-01",
            "budget": 80000,
            "spentBudget": 75000,
            "teamLead": 105,
            "assignedEmployees": [105, 103],
            "milestones": [
                {
                    "name": "Tool Selection",
                    "dueDate": "2025-03-31",
                    "completed": true
                },
                {
                    "name": "Implementation",
                    "dueDate": "2025-05-31",
                    "completed": true
                },
                {
                    "name": "Training & Launch",
                    "dueDate": "2025-07-01",
                    "completed": true
                }
            ]
        }
    ],

    // Client Data
    "clients": [
        {
            "id": 2001,
            "name": "Acme Corporation",
            "industry": "Manufacturing",
            "contactPerson": "John Smith",
            "email": "john.smith@acme.com",
            "phone": "(555) 123-4567",
            "address": {
                "street": "100 Industrial Blvd",
                "city": "Philadelphia",
                "state": "PA",
                "zipCode": "19102"
            },
            "contractValue": 250000,
            "startDate": "2024-01-15",
            "endDate": "2026-01-15",
            "status": "Active",
            "accountManager": 104
        },
        {
            "id": 2002,
            "name": "Global Retail Inc",
            "industry": "Retail",
            "contactPerson": "Mary Johnson",
            "email": "mary.johnson@globalretail.com",
            "phone": "(555) 987-6543",
            "address": {
                "street": "500 Commerce St",
                "city": "New York",
                "state": "NY",
                "zipCode": "10001"
            },
            "contractValue": 180000,
            "startDate": "2024-06-01",
            "endDate": "2025-12-01",
            "status": "Active",
            "accountManager": 105
        },
        {
            "id": 2003,
            "name": "HealthTech Solutions",
            "industry": "Healthcare",
            "contactPerson": "Dr. Sarah Wilson",
            "email": "sarah.wilson@healthtech.com",
            "phone": "(555) 456-7890",
            "address": {
                "street": "250 Medical Center Dr",
                "city": "Boston",
                "state": "MA",
                "zipCode": "02101"
            },
            "contractValue": 320000,
            "startDate": "2025-02-01",
            "endDate": "2027-02-01",
            "status": "Active",
            "accountManager": 104
        }
    ],

    // Financial Data
    "financials": {
        "quarterlyRevenue": [
            {
                "quarter": "Q1 2025",
                "revenue": 425000,
                "expenses": 380000,
                "profit": 45000
            },
            {
                "quarter": "Q2 2025",
                "revenue": 480000,
                "expenses": 420000,
                "profit": 60000
            },
            {
                "quarter": "Q3 2025",
                "revenue": 510000,
                "expenses": 450000,
                "profit": 60000
            }
        ],
        "annualTargets": {
            "revenue": 2100000,
            "profit": 315000,
            "newClients": 8,
            "employeeGrowth": 5
        }
    },

    // System Settings
    "settings": {
        "timezone": "America/New_York",
        "currency": "USD",
        "dateFormat": "MM/DD/YYYY",
        "workingHours": {
            "start": "09:00",
            "end": "17:00"
        },
        "holidays": [
            "2025-01-01", // New Year's Day
            "2025-07-04", // Independence Day
            "2025-11-28", // Thanksgiving
            "2025-12-25"  // Christmas
        ]
    }
};

// Complete DataUtils object with all utility functions
const DataUtils = {
    // Employee utilities
    getEmployeeById: function(id) {
        return jsonData.employees.find(emp => emp.id === id);
    },

    getEmployeesByDepartment: function(department) {
        return jsonData.employees.filter(emp => emp.department === department);
    },

    getActiveEmployees: function() {
        return jsonData.employees.filter(emp => emp.isActive);
    },

    getEmployeesByManager: function(managerId) {
        return jsonData.employees.filter(emp => emp.manager === managerId);
    },

    // Department utilities
    getDepartmentById: function(id) {
        return jsonData.departments.find(dept => dept.id === id);
    },

    getDepartmentByName: function(name) {
        return jsonData.departments.find(dept => dept.name === name);
    },

    getDepartmentSalaryTotal: function(department) {
        const employees = this.getEmployeesByDepartment(department);
        return employees.reduce((total, emp) => total + emp.salary, 0);
    },

    // Project utilities
    getProjectById: function(id) {
        return jsonData.projects.find(project => project.id === id);
    },

    getProjectsByStatus: function(status) {
        return jsonData.projects.filter(project => project.status === status);
    },

    getProjectTeam: function(projectId) {
        const project = this.getProjectById(projectId);
        if (!project) return [];
        return project.assignedEmployees.map(empId => this.getEmployeeById(empId));
    },

    getProjectProgress: function(projectId) {
        const project = this.getProjectById(projectId);
        if (!project) return 0;
        
        const completedMilestones = project.milestones.filter(m => m.completed).length;
        return Math.round((completedMilestones / project.milestones.length) * 100);
    },

    // Client utilities
    getClientById: function(id) {
        return jsonData.clients.find(client => client.id === id);
    },

    getActiveClients: function() {
        return jsonData.clients.filter(client => client.status === 'Active');
    },

    getTotalContractValue: function() {
        return jsonData.clients.reduce((total, client) => total + client.contractValue, 0);
    },

    // Financial utilities
    getCurrentQuarterRevenue: function() {
        const current = jsonData.financials.quarterlyRevenue;
        return current[current.length - 1];
    },

    getYearToDateRevenue: function() {
        return jsonData.financials.quarterlyRevenue.reduce((total, quarter) => total + quarter.revenue, 0);
    },

    getYearToDateProfit: function() {
        return jsonData.financials.quarterlyRevenue.reduce((total, quarter) => total + quarter.profit, 0);
    },

    // General utilities
    formatCurrency: function(amount) {
        return new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: 'USD'
        }).format(amount);
    },

    formatDate: function(dateString) {
        const date = new Date(dateString);
        return date.toLocaleDateString('en-US');
    },

    calculateAge: function(dateString) {
        const today = new Date();
        const birthDate = new Date(dateString);
        let age = today.getFullYear() - birthDate.getFullYear();
        const monthDiff = today.getMonth() - birthDate.getMonth();
        if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
            age--;
        }
        return age;
    },

    // Search functionality
    searchEmployees: function(searchTerm) {
        const term = searchTerm.toLowerCase();
        return jsonData.employees.filter(emp => 
            emp.fullName.toLowerCase().includes(term) ||
            emp.email.toLowerCase().includes(term) ||
            emp.position.toLowerCase().includes(term) ||
            emp.department.toLowerCase().includes(term)
        );
    },

    searchProjects: function(searchTerm) {
        const term = searchTerm.toLowerCase();
        return jsonData.projects.filter(project => 
            project.name.toLowerCase().includes(term) ||
            project.description.toLowerCase().includes(term) ||
            project.status.toLowerCase().includes(term)
        );
    }
};

// Export for use in other modules (if needed)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { jsonData, DataUtils };
}
```

## Conclusion

The embedded JavaScript approach provides the most reliable solution for integrating JSON data with HTML pages in local file scenarios, particularly when using MS Access with the Edge browser control. This method eliminates common issues with CORS restrictions while providing immediate data availability and excellent performance.

The comprehensive data structure and utility functions demonstrated here provide a solid foundation for building interactive business applications with rich data management capabilities.