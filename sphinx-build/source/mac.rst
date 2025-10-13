===
Mac
===

This guide explains how to install the **IBM Envizi Emissions API** add-in from an XML manifest file on
**macOS**.

1. Open `Finder`
2. On the top menu bar, click `Go`
3. From the drop down menu that appears, click `Go to folder`

.. code-block:: none

   /Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents
   
.. note::

   - <username> should be replaced with the name of your Mac user

4. If the `wef` folder does not exist, right-click and create a folder called `wef`
5. Click on the `wef` folder to enter it
6. Place your downloaded `manifest.xml` file in the `wef` folder

.. image:: _images/placing-wef-file.png
   :alt: Placing manifest file in Mac directory
   :align: center
   
After successfully placing the manifest file, the add-in will appear in the Developer Add-in section.

.. image:: _images/developer-add-in.png
   :alt: Developer Add-in in Excel
   :align: center


For more information please see the `Microsoft 365 Office Add-in Mac <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac>`_ documentation for Mac.

Troubleshooting
---------------

- **Add-in not appearing**  
   - Make sure the manifest file is placed in the correct directory
   - Ensure Excel has been restarted after placing the manifest file

Next Steps
----------

After installation, please refer to the :doc:`Calculation Mode Tip </installation>` section for optimizing your Excel configuration.

Please follow :doc:`Usage </use>` documentation for next steps.
