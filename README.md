# Strava Segment Leaderboard

[![Latest Release](https://img.shields.io/github/v/release/eric-benny/segment-leaderboard)](https://github.com/eric-benny/segment-leaderboard/releases)
[![License](https://img.shields.io/github/license/eric-`benny/segment-leaderboard)](LICENSE)

Chrome extension for extracting shareable monthly top-10 leaderboard cards from Strava segments.

## Features

- Extract monthly top-10 leaderboards from any Strava segment
- Customizable date ranges (This Month, This Week, This Year, All Time)
- Self-contained downloadable HTML output
- Real-time progress tracking during fetch
- Automatic session reuse (no separate authentication needed)
- Supports both male and female leaderboards

## Installation

This Chrome extension is distributed for local developer mode installation only.

### Prerequisites
- Google Chrome (or Chromium-based browser like Edge, Brave, etc.)

### Steps

1. **Download the Extension**
   - Go to the [Releases page](../../releases)
   - Download the latest `strava-segment-leaderboard-v{version}.zip` file
   - Unzip to a permanent folder location (do not delete after installation)

2. **Install in Chrome**
   - Open Chrome and navigate to `chrome://extensions/`
   - Enable **Developer mode** using the toggle in the top-right corner
   - Click **Load unpacked**
   - Select the unzipped folder containing `manifest.json`
   - The extension should now appear in your extensions list

3. **Verify Installation**
   - Navigate to any Strava segment page (e.g., `https://www.strava.com/segments/12345`)
   - You should see a "Monthly Leaderboard" button injected on the page
   - Click the extension icon in your Chrome toolbar to access the popup interface

### Updating the Extension

When a new version is released:
1. Download the new zip file from Releases
2. Unzip to the same folder (overwrite existing files)
3. Go to `chrome://extensions/`
4. Click the refresh icon on the extension card

## Usage

1. Navigate to any Strava segment page
2. Click the "Monthly Leaderboard" button OR use the popup to enter segment ID
3. Select date range (This Month, This Week, This Year, All Time)
4. Extension fetches leaderboard data and generates a downloadable HTML card
5. Save and share the top-10 leaderboard card

## Development

This extension uses vanilla JavaScript with no build process.

### Local Development
1. Clone this repository
2. Make changes directly in the `src/` directory
3. Load the `src/` folder as an unpacked extension in Chrome
4. Reload the extension after changes

### Release Process
- Merges to `main` branch automatically trigger a new release
- Version is auto-incremented in `manifest.json`
- GitHub Release created with zip file attached
