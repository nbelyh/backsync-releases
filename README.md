# Visio BackSync Add-in

The Visio BackSync Add-in enables bidirectional data synchronization within Microsoft Visio, allowing you to update external data sources directly from Visio. It seamlessly integrates with Visio's data-binding infrastructure and is compatible with various data sources such as Excel files, databases, and SharePoint lists.

![Visio BackSync Screenshot](https://unmanagedvisio.com/wp-content/uploads/24wy6mu.png)

## Download and Installation

- [Download Latest Version (1.1.2)](https://unmanagedvisio.com/download/back_sync/BackSync-1.1.2.msi)
- The add-in supports Visio 2007 SP3, 2010, 2013, 2016, 2019, and Visio Plan 2 for both x86 and x64 versions that support shape data linking.

## Features

- **Bidirectional Data Synchronization:** Supports create, update, and delete operations in Visio with external data sources.
- **Shape Text and Properties Synchronization:** Syncs shape text and properties such as position, size, color, and more.
- **Action Preview:** Allows you to preview changes before applying them, with options to exclude, revert, or modify actions.
- **Custom Attribute Synchronization:** Easy setup for attributes like position, size, and color.
- **Conflict Resolution:** Offers a resolution window for conflicting changes.
- **Programmatic Access:** Supports calling the add-in from VBA.

![Preview Changes](https://unmanagedvisio.com/wp-content/uploads/06-12-2016-04-42-52.png)

## How it Works

The add-in queries data from the external source and analyzes changes in both the data source and Visio diagram. If conflicts arise, the add-in provides a resolution interface. For each data connection, a dialog allows synchronization management using a simple button added to the Data panel.

## Data Source Compatibility

Should be Compatible with any valid Visio data source supporting updates, tested with Excel, SQL, and SharePoint lists.
