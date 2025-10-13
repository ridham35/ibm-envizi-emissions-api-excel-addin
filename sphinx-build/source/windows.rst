=======
Windows
=======

This guide explains how to install an **IBM Envizi Emissions API** add-in from an XML manifest file on
**Windows** using the *Trusted Add-in Catalogs* method.

.. important::
   Requires Microsoft Excel 2021 or later

Step 1: Locate Your Manifest File
---------------------------------

1. Make sure you already have the ``manifest.xml`` file
2. Save this file in an easy-to-access folder, such as: ``C:\Addins\``

.. image:: _images/placing-file.png
   :alt: Placing the file
   :align: center


Step 2: Share the Folder
------------------------

Make sure your file is placed on Network path or in Shared Folder.

1. To share the folder, right-click on the folder and select **Properties**.
2. Navigate to the **Sharing** tab.
3. Click on the **Share** button.
4. Add New/Existing user with appropriate access permissions.
5. After submitting, you will see a **Network Path** as follows: (keep this path handy)

   .. code-block:: none

      \\<user>\Addins

.. image:: _images/share-folder.png
   :alt: Shared folder properties
   :align: center
   
Step 3: Open Excel and Access Options
-------------------------------------

1. Launch **Microsoft Excel**.
2. Go to the **File** menu (top-left corner).
3. Scroll down and click **Options** (bottom-left of the menu).
4. A window called **Excel Options** will appear.

.. image:: _images/options.png
   :alt: Options dialog box
   :align: center

Step 4: Open the Trust Center Settings
--------------------------------------

5. In the **Excel Options** window, look at the left-side menu.
6. Click **Trust Center** (last item in the list).
7. On the right-hand side, click the button **Trust Center Settings...**.

.. image:: _images/trust-center-1.png
   :alt: Trust Center settings
   :align: center

Step 5: Configure Trusted Add-in Catalogs
----------------------------------------------

1. In the new **Trust Center** window, select **Trusted Add-in Catalogs** from the left menu.
2. In the **Catalog URL** box, type the full Network Path to your manifest folder.
   Example:

   ``\\<user>\Addins``

   .. important::
      Do not type the XML file path here.
      Only provide the **Network Path** from the sharing properties.

3. After entering the Network Path, click **Add Catalog**.
4. The Network Path will now appear in the list below.
5. Select the checkbox **Show in Menu** so the add-in will be visible in Excel.
6. Click **OK** to save your changes.
7. Click **OK** again to close Excel Options.

.. image:: _images/trust-center-2.png
   :alt: Trust Center catalog configuration
   :align: center

Step 6: Restart Excel
---------------------

1. Close all open Excel windows.
2. Open Excel again to refresh the add-ins.

Step 7: Open Your Add-in in Excel
---------------------------------

1. Go to the **Insert** tab in Excel.
2. Click **My Add-ins** in the toolbar.

.. image:: _images/more-add-in.png
   :alt: Add-ins dialog
   :align: center

3. A dialog box will appear. Look for the section named **Shared Folder**.
4. **IBM Envizi Emissions API** add-in should now be listed here.
5. Select the add-in and click **Add**.

.. image:: _images/installed-windows.png
   :alt: Add-In installed
   :align: center

Step 8: Using the Add-in
------------------------

- Once installed, the Welcome task pane will appear and add-in is ready to use.

.. image:: _images/welcome.png
   :alt: Welcome Page
   :align: center

Troubleshooting
---------------

- **Add-in not appearing**  
   - Make sure manifest file is on network path or shared folder.
   - Ensure the manifest network path is correct and listed in Trusted Add-in Catalogs.
   - Ensure you restarted Excel after saving settings.

- **Still not working**
   - Delete the catalog path, re-add it, and restart Excel again.
   - Confirm you are running **Office 2021 or later** (older versions may not support this method).

- **Multiple add-ins**
   - If you have multiple manifest files, we recommend placing them in separate folders to avoid conflicts.

Next Steps
----------

After installation, please refer to the :doc:`Calculation Mode Tip </installation>` section for optimizing your Excel configuration.

Please follow :doc:`Usage </use>` documentation for next steps.

