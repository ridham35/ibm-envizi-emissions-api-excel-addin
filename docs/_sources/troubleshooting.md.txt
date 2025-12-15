# Troubleshooting Guide

## Overview

This document provides a comprehensive guide to common error types that can be thrown by IBM Envizi - Emissions Calculations in Excel, along with troubleshooting steps to resolve them.

---

## Table of Contents

1. [Plugin Specific Errors](#plugin-specific-errors)
2. [Errors Thrown from SDK Calls](#errors-thrown-from-sdk-calls)

---

## Plugin Specific Errors

### Value Not Available Error

#NA gets populated in the cell after calling the ENVIZI excel function

**Causes:**

- Authentication to the IBM Envizi Emission API not done before calling the function to calculate errros


**Troubleshooting Steps:**

- Fill in the API key, Tenant ID and Org ID and then click on Login to authenticate.
- Call the appropriate Envizi function

### Invalid credentials Error

**Causes:**

Account Credentials entered is incorrect

**Troubleshooting Steps:**

- Make sure the API Key, Tenant ID and Org ID are copied appropriately from the Overview Dashboard into the 'Account credentials' section within the Excel template
- Save credentials and Login

### Alert in case of invalid cell values

An alert popup that says 'This value doesn’t match the data validation restrictions defined for this cell.'

**Causes:**

Unaccepted values populated in the cells used to make calculation

**Troubleshooting Steps:**

Refer to the dropdown and choose only one of the accepted and valid values for the specific cell

## Errors Thrown from SDK Calls

### Error in Value

#VALUE! gets populated in the cell after calling the ENVIZI function

**Causes:**

- Invalid data type could be set
- Invalid location (country and/or state province is wrongly set)
- Invalid unit is set

**Troubleshooting Steps:**

- Hover over the yellow exclamation mark that appears next to '#VALUE!' to understand the actual error that is causing the issue
- Make sure to check that each field is validated correctly before making the call 