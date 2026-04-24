# Utility Suite 🚀

A premium, high-performance suite of browser-based utilities built with **React**, **Tailwind CSS v4**, and **Vite**. Designed to streamline administrative workflows and data validation.

## ✨ Features

### 1. Duplicate Tickets Validator
The flagship tool in the suite, designed to cross-verify order reports against ticket listings with complex matching logic.
- **Single/Dual Mode**: Process a combined Excel sheet or two separate reports (Orders vs Tickets).
- **Intelligent Logic**: 
  - **Single Tickets**: Automated 1:1 matching.
  - **Return Tickets**: Automated 1:2 matching (validates double appearances).
- **Smart Column Detection**: Automatically identifies headers like `pg_order_id`, `num_of_tickets`, and `order_id`.
- **Export**: Instantly download mismatched records into a clean Excel file.

## 🛠️ Technology Stack
- **Frontend**: React 19 (Vite)
- **Styling**: Tailwind CSS v4 (Modern CSS-first engine)
- **Icons**: Lucide React
- **Excel Engine**: SheetJS (XLSX)

## 🚀 Getting Started

### Prerequisites
- [Node.js](https://nodejs.org/) (v18 or higher)
- npm or yarn

### Installation
1. Clone the repository:
   ```bash
   git clone <your-repo-url>
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Start the development server:
   ```bash
   npm run dev
   ```

## 📁 Project Structure
- `src/pages/Home.jsx`: The central hub for all utilities.
- `src/pages/TicketValidator.jsx`: Core logic for the ticket validation tool.
- `src/index.css`: Global styles and Tailwind v4 configuration.

## 📄 License
This project is private and intended for internal administrative use.

---
Built with ❤️ for better workflows.
