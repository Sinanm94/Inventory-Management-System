# Inventory Management System

## Overview
This Google Apps Script project serves as an inventory management system, allowing users to perform various transactions such as receipts, shipments, transfers, and inventory adjustments.

## Features
- **User-friendly Menu System:** Interactive menu sheets for different transaction types.
- **Dropdown Menus:** Dropdown menus for selecting parts, locations, and quantity, ensuring data integrity.
- **Real-time Updates:** Automatic updates to the inventory sheet based on user transactions.
- **Data Validation:** Validation checks to ensure accurate and valid data entry.

## Usage
1. Open the Google Sheets document.
2. Use the provided menu sheets to perform transactions (Receipt, Shipment, Transfer, Inventory Adjustment).
3. Follow the on-screen instructions for each transaction type.

## Functions

### `myReceiptView()`
Clears data and sets up the menu sheet for Receipt transactions.

### `myShipmentView()`
Clears data and sets up the menu sheet for Shipment transactions.

### `myTransferView()`
Clears data and sets up the menu sheet for Transfer transactions.

### `myInvAdjView()`
Clears data and sets up the menu sheet for Inventory Adjustment transactions.

### `submitResults()`
Processes user input from the menu sheet and updates the inventory accordingly.

### `onEdit(e)`
Triggers when a user edits the spreadsheet, providing dynamic information based on user actions.

## Contributing
Contributions are welcome! Please follow our [contribution guidelines](CONTRIBUTING.md).

## License
This project is licensed under the [MIT License](LICENSE).

## Acknowledgements
- Thanks to [Google Apps Script](https://developers.google.com/apps-script) for making automation in Google Workspace possible.

## Contact
For inquiries, contact [Muhammed Sinan](mailto:muhammedsinan203@gmail.com).
