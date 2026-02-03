# Annual Lottery System User Guide (English)

## Features

- Pure front-end: No server required. Open the HTML file and use it directly.
- Local storage: All data is stored in the browser and persists after refresh.
- Excel import: Bulk import for employee and boss lists.
- Manual management: Add, edit, and delete participants.
- Fair randomness: Uses the native random algorithm for unbiased draws.
- Boss privilege: Bosses can stay in the pool and win repeatedly.
- Modern UI: Dynamic effects with a particle background.

## File Structure

```
lotteryCustom/
├── lottery.html           # Main lottery page
├── Usage_Guide_EN.md      # English user guide
└── 抽奖功能说明.md          # Feature requirements (Chinese)
```

## Quick Start

1. Open `lottery.html` in a modern browser.
2. Go to the "Manage" tab and import or add participants.
3. Go to the "Lottery" tab and start drawing.

## Detailed Steps

### 1. Import Participants

#### Excel format requirements
- Supports `.xlsx` and `.xls`
- The first row must be the header: `Name`, `Number`
- Example:
  | Name  | Number |
  |-------|--------|
  | Alice | 001    |
  | Bob   | 002    |
  | Carol | 003    |

#### Import flow
1. Click the "Manage" tab.
2. Choose an Excel file for Employees or Bosses.
3. The system imports and de-duplicates automatically.

### 2. Add Participants Manually

1. In the "Manage" tab, find the "Add Manually" section.
2. Fill in name and number.
3. Select participant type (Employee/Boss).
4. Click "Save".

### 3. Start a Draw

1. Click the "Lottery" tab.
2. Click "Start". The sphere spins.
3. Click "Stop" (or wait for auto-stop). The winner appears in the center, and is added to the results list.

### 4. Boss Rules

- Bosses show a special label.
- Bosses stay in the pool after winning and can win again.

### 5. Data Management

- Reset results: Click "Reset" to clear all draw results.
- Clear all data: In "Settings", click "Clear All Data" to delete all participants and results.

## Notes

1. Use a modern browser (Chrome, Firefox, Edge).
2. Data is stored locally in the browser. Back up if needed.
3. Clearing browser data or switching browsers will erase saved data.
4. Avoid refreshing during a draw.

## Technical Details

- UI: Bootstrap 5
- Data processing: SheetJS (xlsx.js)
- Animations: CSS3
- Storage: localStorage

## FAQ

### Q: Excel import does not work. Why?
A: Check the file format and ensure the first row is `Name` and `Number`.

### Q: How is fairness ensured?
A: The system uses JavaScript's native `Math.random()` for unbiased selection.

### Q: Why can bosses win more than once?
A: This is a special rule for scenarios where bosses remain in the pool after winning.

### Q: Can I export data?
A: Export is not supported in the current version.

## Changelog

- v1.0.0 (2024-01-01)
  - Initial release
  - Basic lottery flow
  - Excel import
  - Boss rule support
  - Modern UI
