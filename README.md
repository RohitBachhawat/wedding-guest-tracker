# 🎊 Wedding Guest Tracker & Logistics Coordinator

A high-performance, AI-powered web application designed for wedding coordinators to manage guest arrivals, departures, and ground transportation with surgical precision.

![Version](https://img.shields.io/badge/version-1.0.0-gold)
![License](https://img.shields.io/badge/license-MIT-blue)

## ✨ Key Features

- **🤖 AI Ticket Extraction**: Upload IRCTC/Flight PDFs or screenshots; Google Gemini automatically extracts names, PNRs, times, and hubs.
- **🚗 Smart Pooling Engine**: Automatically suggests guest groups that can share a vehicle based on arrival times, locations, and car capacity.
- **📡 Live Status Tracking**: One-tap flight and train status updates powered by AI to stay ahead of delays.
- **⚡ Performance Optimized**: 
    - **Optimistic UI**: Instant status toggles with background sync.
    - **Offline-First**: Works without internet using local caching and a robust sync queue.
    - **Zero-Wait Boot**: Concurrent data fetching for instant app readiness.
- **📊 Google Sheets Backend**: Uses Google Apps Script as a lightweight database for easy multi-user collaboration and data export.

## 🛠️ Tech Stack

- **Frontend**: Vanilla HTML5, CSS3 (Modern Flexbox/Grid), JavaScript (ES6+).
- **AI Engine**: Google Gemini Flash API.
- **Backend/Database**: Google Apps Script (GAS) + Google Sheets.
- **Aesthetics**: Premium "Royal Wedding" theme with glassmorphism and micro-animations.

## 🚀 Setup Instructions

### 1. Backend Setup (Google Sheets)
1. Create a new Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Paste the code from your App Script file.
4. Click **Deploy > New Deployment**.
5. Select **Type: Web App**, set **Execute as: Me**, and **Who has access: Anyone**.
6. Copy the **Web App URL**.

### 2. Frontend Configuration
1. Open `index_wg.html` in any browser.
2. Enter your **Google Apps Script URL** and **Gemini API Key** in the setup screen.
3. Define your wedding city and logistics hubs (Airports/Stations).
4. Start tracking!

## 📱 Mobile Experience
The app is built as a **Progressive-style Web App**, optimized for one-handed use by ground coordinators at busy airports and railway stations.

---
*Created with ❤️ for seamless wedding logistics.*
