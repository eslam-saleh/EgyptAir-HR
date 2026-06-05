# Employee Holiday Planner

An interactive calendar for planning employee personal holidays against Egyptian official holidays and the employee's annual balance.

## Features

- Full current-year calendar from January 1 through December 31.
- User-controlled annual holiday balance using a list or custom counter.
- Personal holiday selection for valid working days only.
- Blocks personal holidays on official Egyptian holidays.
- Blocks personal holidays on weekends in Egypt: Friday and Saturday.
- Live selected-days and remaining-balance counters.
- Live employee leave summary in a readonly text area.
- English and Arabic versions with right-to-left layout for Arabic.
- Browser storage saves selections separately for each calendar year.

## How to Use

1. Open `index.html` in a browser, or serve this folder locally.
2. Choose the employee's annual balance from the list or type a custom number.
3. Click a valid working day to mark it as a personal holiday.
4. Click a selected day again to remove it.
5. Use the language buttons to switch between English and Arabic.
6. Use the summary text area at the bottom to review the employee's selected days.

## Selection Rules

- The calendar always covers the current year.
- Friday and Saturday cannot be selected.
- Official Egyptian holidays cannot be selected.
- The employee cannot select more personal holidays than the annual balance.
- The reset button clears the current year's selections.

## File Structure

```text
Web_Assignment/
|-- index.html
|-- styles.css
|-- script.js
`-- README.md
```

## Technologies Used

- HTML5
- CSS3
- Vanilla JavaScript
