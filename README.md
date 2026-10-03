# Sales Performance Dashboard

A production-quality Business Intelligence Dashboard built with Flask, SQLite, Bootstrap 5, and Chart.js. Allows users to analyze sales performance, monitor KPIs, identify business trends, and make data-driven decisions.

## Features

- **Login System** with session management and password hashing
- **KPI Cards**: Total Sales, Revenue, Profit, Orders, Customers, Profit Margin, AOV, Today's Sales
- **12 Charts**: Monthly trends, category/regional breakdowns, product rankings, scatter plots, distributions
- **Advanced Filters**: Year, Month, Date Range, Region, State, City, Category, Segment, Product Search
- **Interactive Tables**: Products, Customers, Orders with search, sort, and pagination
- **Reports**: Export to CSV, Excel (XLSX), PDF, and Print
- **Analytics Page**: Automated business insights and recommendations
- **Dark Mode Toggle**
- **Responsive Collapsible Sidebar**
- **Animated KPI Counters**
- **Loading Skeletons and Spinners**
- **REST API** endpoints for all dashboard data

## Tech Stack

| Layer    | Technology                   |
|----------|------------------------------|
| Backend  | Python 3, Flask              |
| Database | SQLite                       |
| Frontend | HTML5, CSS3, Bootstrap 5     |
| Charts   | Chart.js                     |
| Reports  | openpyxl, reportlab          |

## Installation

```bash
# Clone the repository
git clone <repo-url>
cd SalesDashboard

# Install dependencies
pip install -r requirements.txt

# Create the database
python database/create_database.py

# Seed sample data (500+ orders)
python database/seed_data.py

# Run the application
python app.py
```

Open **http://127.0.0.1:5000** in your browser.

## Default Login

| Username | Password  |
|----------|-----------|
| admin    | admin123  |

## Project Structure

```
SalesDashboard/
├── app.py
├── requirements.txt
├── README.md
├── sales.db
├── database/
│   ├── __init__.py
│   ├── create_database.py
│   └── seed_data.py
├── templates/
│   ├── base.html
│   ├── login.html
│   ├── dashboard.html
│   ├── products.html
│   ├── customers.html
│   ├── reports.html
│   └── analytics.html
└── static/
    ├── css/
    │   └── style.css
    ├── js/
    │   └── dashboard.js
    ├── images/
    └── icons/
```

## API Endpoints

| Endpoint             | Method | Description               |
|----------------------|--------|---------------------------|
| /api/dashboard       | GET    | KPI summary               |
| /api/monthly-sales   | GET    | Monthly sales & profit    |
| /api/category-sales  | GET    | Sales by category         |
| /api/region-sales    | GET    | Sales by region           |
| /api/state-sales     | GET    | Sales by state            |
| /api/customer-growth | GET    | Segment distribution      |
| /api/top-products    | GET    | Top/bottom products       |
| /api/orders          | GET    | Order listing             |
| /api/products        | GET    | Product listing           |
| /api/customers       | GET    | Customer listing          |
| /api/analytics       | GET    | Business insights         |

All API routes accept filter query parameters: `year`, `month`, `region`, `state`, `city`, `category`, `sub_category`, `segment`, `search`.

## Screenshots

> Screenshots placeholder - add images here after running the app.

## Future Improvements

- User registration and role-based access control
- Real-time data updates with WebSockets
- More chart types (radar, treemap, funnel)
- Email report scheduling
- Data import from CSV/Excel
- Multi-language support
- Unit and integration tests
