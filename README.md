# Excel - Gantt Chart for Project Management

</br>

![yhtyhj](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/05f08dbb-a8a9-4957-9eff-1acb47086607)

## Index

- [Overview](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/tree/main#overview)

- [User Guide](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/tree/main#user-guide)

- [Documentation](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/tree/main#documentation)
  - [Initial design setup (Gantt & Settings Worksheets)](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#1-initial-design-setup-gantt--settings-worksheets)
  - [Basic item details with Dynamic Structuring](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#2-basic-item-details-with-dynamic-structuring)
  - [Role & Team Management system](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#3-role--team-management-system)
  - [Manual Planning with independent Start or End Dates](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#4-manual-planning-with-independent-start-or-end-dates)
  - [Dynamic Planning with Dependency Engine](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#5-dynamic-planning-with-dependency-engine)
  - [Dynamic Project Timespan](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#6-dynamic-project-timespan)
  - [Auto-Coloring Engine with 4 automated color modes](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#7-auto-coloring-engine-with-4-automated-color-modes)
  - [Base Plan Comparison](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#8-base-plan-comparison)
  - [Progress Tracking](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#9-progress-tracking)
  - [Date Highlighting (Dynamic & Static)](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#10-date-highlighting-dynamic--static)
  - [Scroll Buttons](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#11-scroll-buttons)
  - [Manage & Filter Items (Via Quick Access Toolbar)](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management#12-manage--filter-items-via-quick-access-toolbar)


</br>
</br>

## Overview

<p align="justify"> Thank you for visiting this repository. The primary aim of this resource is to provide guidance on creating an Excel Gantt chart for project management. The attached chart is a useful tool for planning and organizing tasks in a project. With Excel, you can create both basic and complex versions of the Gantt chart. Our template incorporates advanced and interactive features inspired by the top project management software, Jira and Trello. It is a professional tool that any team can use to effectively plan and manage their own projects. </p>

<p align="justify"> This template introduces a novel and intuitive way of creating and managing fully dynamic project schedules in Excel with items depending on each other based on the four common dependency connection types: Finish-to-Start (FS), Start-to-Start (SS), Start-to-Finish (SF), or Finish-to-Finish (FF). Furthermore, the template includes a smart Project Role & Team Management system, a fully automated Colouring Engine with four different color modes, Plan vs. Base comparison, and more amazing features. It is designed using a fundamentally new approach that allows for a tremendous degree of automation, interactivity, and visualization of different project perspectives. </p>

Please be aware that this repository's visual explanations are presented through gifs. In case you encounter issues with the size, simply click on the gif and a new tab will open, allowing you to increase its size by using Ctrl + Scroll wheel mouse up.

[Link to Excel Workbook](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/blob/main/Gantt%20Chart%20for%20Project%20Management.xlsx)

- Unicode icons: ⚠ - ◆ - ⚑ - ↖ - ↙
- Color palette of the workbook:  F0E936 - 40D492 - F9B710 - FE1684 - 9237BC - 4633F2 - 3285F3 - 1DD7F3

</br>
</br>

## User Guide

<p align="justify"> Our final Excel template allows you to set up an advanced and workday-focused project schedule consisting of project stages with task and milestone items. Setting up a schedule for these items is as easy as defining the number of required workdays for each item and then either manually entering an independent start or end date or make use of the core feature of this template, a fully implemented dependency engine. </p>

<p align="justify"> The dependency engine allows you to make one item dependent on another by creating a connection through an intuitive ID system. The columns next to the referenced ID instantly display all the fundamental information about the referenced item. There are four possible dependency connection types to choose from, such as a forward scheduling finish-to-start (FS) or start-to-start dependency (SS), or a backward scheduling start-to-finish (SF) or finish-to-finish dependency (FF). </p>

![ethjh](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/fc8e72a3-cd9f-47dd-afdc-981ea1244db5)

<p align="justify"> Notice how both the stage bar and the project time span instantly adapt to the changes made. The way this dependency engine is built enables you to set up consecutive dependencies with different dependency connection types in seconds. Any new task or milestone item you add below a stage is instantly assigned to that stage via an automated stage ID system. </p>

</br>
</br>

<p align="justify"> In the same way, you can quickly create a new stage with its own items. Notice how the whole area and the item names are automatically formatted and structured based on the selection you make. With this dependency engine, you are not limited to creating dependencies within a stage. You can also make stages dependent on other stages. For a forward-oriented stage dependency, you only have to make the initial item in your stage dependent on another stage. Then, a set of standardized formulas allows you to make the stage update based on its items, and finally, you can make the other items consecutively dependent on that initial item. Define the number of workdays and make sure that the milestone ends together with the preceding task. </p>

![UG_2](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/365375e5-8c29-4a1e-96cd-70f1e55b69e3)

</br>
</br>

<p align="justify"> In addition, you might have noticed that every stage automatically gets colored with one color group per stage and different color shades for the stage, the tasks, and the milestone symbol. This is possible through an automated coloring engine that dynamically creates and applies color codes in the background. </p>

![UG_3](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/e72a66a5-93b3-4632-9e59-80d66ed36ba8)

<p align="justify"> The team roles color mode allows you to visualize the project based on the respective role that the project team members responsible for an item got assigned. When we open the "Settings" worksheet we get access to an incredibly powerful project role and team management system that allows us to add an unlimited amount of team members and group them into up to eight roles. You only need to add the role name and provide a short code in the "Project Role Types" table. Then you can add new team members and assign any of the roles in the "Project Team" table. The assignment of roles is automatically tracked in the count column of the role type section. Once the roles are defined and the team members added and assigned, jump back to the "Gantt Chart" worksheet and you can assign any of the defined team members to your tasks and milestones. The respective role is then automatically displayed and visualized with its own individual color in the Gantt chart, the coloring isn't linked to the team member but to its role. </p>

![UG_4](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/c8f998f8-b40c-46f1-bf95-ff9f9e579103)

<p align="justify"> In the last color mode, the issues color mode, every item that gets marked with an issue symbol is highlighted in a warning red color. Of course, you can always deactivate all the coloring by switching to "Default"; however, the normal coloring in every Gantt chart is the "Project Structure" color mode. </p>

![UG_5](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/e2495135-326c-4feb-b2de-038f96942a6e)

</br>
</br>

<p align="justify"> Once you have set up your project schedule, you might want to keep track of the adjustments that you make to that schedule over time. You can document the date of the latest update to the actual project plan and then, in order to preserve a snapshot of that base plan, open the base plan section (the collapsed section with dark blue headers). Taking a snapshot is as simple as copying as values the start and end dates of the actual plan section, and then paste it into the base plan section. When time passes by and you make change to the actual plan, you now have the original plan and the actual plan side by side. In addition, you can visualize the plan shifted dates by changing the "Display" mode to "Plan vs Base". </p>

![UG_6](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/7b8ff7ea-4f87-4471-bdf1-90bd023f2a71)

</br>
</br>

<p align="justify"> This template also allows you to keep track of the stage, task, and milestone completion. Once you set the "Show progress" mode to yes the percentage completed is beautifully visualized as a filled bar for the stages and tasks and as a finish flag symbol for completed milestones. You can even make the completion of milestones conditionally based on the completion of one or multiple tasks so that the milestone is automatically completed not before the linked task has reached the full 100 percent. In addition, a standardized formula allows you to automatically update the stage progress based on its items. Eventually, this completion tracking feature is perfectly complemented by the date highlighting option which allows you to either dynamically highlight today and make overdue items before that line instantly catches your eyes, or alternatively, you can statically highlight any of the visible dates in the timeline. The selectable timeline dates are immediately updated in the drop-down list whenever you scroll the whole Gantt chart area to the right or left which might be especially helpful for projects of longer duration. </p>

![UG_7](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/5211586a-74ac-43ba-be0b-0d4680095695)

</br>
</br>

# Documentation

## 1. Initial design setup (Gantt & Settings Worksheets)

### 1.1. General design

<p align="justify"> To begin with, let's set up the general design for the Gantt and Settings worksheets. Beginning with the Gantt worksheet, we will start by setting up the header rows for our input and Gantt chart area. We will select these two rows and decrease the general text style to size 10, make it bold, and align it to the center and middle. After that, we will select the range B13 to W13 and give the cells a green fill, as this is the row where we will put all the single-column headers. In the row above, we will create section names starting with the first three columns as the section for ids and then the next six columns for the item details. Column K to T is where the actual planning will happen, while the two following columns will be used to preserve a snapshot of the actual plan at a certain point in time. We will call this base plan. </p>

![Doc1 1](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/3458da0c-6fa9-4f4c-9d04-5ff8b7ae77a5)

<p align="justify"> For both these planning sections, we want to be able to document the dates of an update or the saving of a snapshot. We will make use of these two additional rows here, make the font for the input field a bit smaller and just insert a placeholder date value for now. Same for the base plan section. Eventually, we will add some light gray borders to the section cells. For the single-column headers, we will set the general font color to white and increase the rows' height. We will start by naming the first columns as ID, then "s id" for stage ID, and "r id" for role ID. All of these columns in the ID section will be automatically calculated, so we will color them according to the legend present above, which is a lighter green fill. </p>

![Doc1 2](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/65685e69-da02-4fea-8f1f-1a18419af9ee)

<p align="justify"> Moving to the item details section, the columns for the item type and description are the first actual input columns. The next column will be a special column that only contains the indicator colors, so let's give that one an even brighter green fill. The next one is the column for the project roles, which will be automatically computed based on the team member selected in the following column. The final column in this section makes use of the Unicode icons prepared above, and this one will be the one used to highlight issues in any of your items in the actual plan section. </p>

![Doc1 3](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/cae077dd-3ef3-4eae-8cd5-57499f5f1b6a)

<p align="justify"> In the actual plan section, the most crucial columns are the last two ("Start" and "End"), as these will be the calculated start and end dates for every single item. These two columns will contain the most crucial formulas in this worksheet, and the start and end date calculation will always be based on two types of input. The first one is the number of required workdays for an item, and the second one is some sort of date. It can be either a manually entered independent start date or an independent end date ("ind.start" and "ind.end"), or, and this is where the main power of this template will come from, a dependency connection to another item. This dependency connection will require a link to the other item's ID ("d.id") based on which some crucial information will be displayed in the following two merged columns ("d.item"). Then, we have one column to select the type of dependency connection ("d.conn") and another one if you want to potentially add a lag by which you want the item to be shifted forward or backward ("d.lag"). We group together these state-related input columns from the actual plan section and make them collapsible. Finally, let's color the snapshot columns from the base plan section that will be used to preserve the start and end date and make them collapsible too. The last column will be used for the project tracking, so let's put a percentage symbol in there for now. This is the perfect time to quickly adjust the size of the columns to match the expected size of the input. Right after that, we can start creating the timeline area. </p>

![Doc1 4](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/e184dac6-7fef-4a6e-800c-159c2a52c057)

</br>
</br>

### 1.2. General timeline design

<p align="justify"> For the timeline, we will fill the cell right next to the percentage column header with a dark blue. Right above that, we are going to select two cells, set the fill to a light blue, make the column really tight, and then merge these three cells together. For the merged cell we will set rotate text up, text alignment to the top, and the font size to 8. Then we can just select the whole column and transfer this specific formatting to as many subsequent columns as we wish using the autofill function. </p>

![Doc1 5](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/f9fc0f27-33f3-4317-9ce3-20e608a8358d)

</br>
</br>

### 1.3. Enhancing the overall design aesthetics

<p align="justify"> That's already it for the basic design of the headers and timeline. For a cleaner design, let's hide the default grid lines and add a way better-looking custom formatting instead. Initially, we want to visually separate the main sections, so let's hold the Ctrl key and select the respective sections, including the header rows and the number of content rows below. With all these selected, we go to the Home tab, open the border menu, and click on more borders. This will open the Format Cells window in which we can now select a border color like this mid-light gray and then specify that we want to add a left and right border. The overall spacing in this sheet is aimed to perfectly fit into the open window, with the actual planning section fully expanded and the base plan section collapsed, so do that and then let's select the whole content area and add some horizontal row-wise borders with an even lighter gray. </p>

![Doc1 6](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/7f5d8188-491e-47dd-931b-d86dd16db20d)

</br>
</br>

### 1.4. First configuration of the "Settings" worksheet

<p align="justify"> We are going to add a settings button on the top-right corner of our window to take us directly to the Settings worksheet. To create the button, we will merge four cells and add a white border, label it as "Settings," and align the text to the center and middle. We will then right-click on the button and select "Link" to create an internal link within the workbook. We will choose the "Settings" worksheet after selecting "Place in this document" and then we will select the target cell where we plan to place the button that will take us back to the Gantt chart. We can also define a screen tip, which is the text that appears when hovering over the button. To make the merged cell area clickable, we will activate the "Wrap Text" option. Clicking on the settings button will redirect us to cell X2 on the Settings worksheet. We will also create a similar button to take us back to the Gantt worksheet. </p>

![Doc1 7](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/c1971df4-f2b4-4b7b-b947-605fead64097)

<p align="justify"> The initial input section that we're going to create in this Settings worksheet will be for some basic project details. We add some grey borders around that section and adjust the column size so that we have two small columns on each side acting as padding and two large columns for creating some input fields. Let's also hide the gridlines on the whole worksheet for a cleaner look and then create three input fields: "Project title", "Project manager" and "Project priority". We select all these newly created labels, make them bold, decrease the font size and also add some indent and a light gray border. Then the cell next to each label are the actual input fields which we highlight by adding a light green fill and a similar font size and indent formatting. </p>

![Doc1 8](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/6aee3fbe-d4c6-4a8c-aa00-9fd408fe5eca)

<p align="justify"> Let's enter some example values, the first two are simple text fields, but the project priority field is a good opportunity to quickly demonstrate how to set up a drop-down list. While having selected the desired cell, go to the "Data" tab, "Data tools" section and click on "Data validation". A window will open, we select "List" in the "Allow" field and then type the values to be selected in the "Source" field. In this case, the values will be: "High, Medium, Low". </p>

![Doc1 9](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/76cc0c3f-c88f-4211-b557-0f968a373d67)

<p align="justify"> For now, we will reference the "Project title" and "Project manager" directly in the Gantt worksheet, so let's replace the current title with the input value from the settings worksheet, make it an absolute reference with the dollar signs, and then let's add some labels in the dark blue header, to the left of the settings button. The first label will be the "Project manager", merge the cells, make it bold and green. The value for this label will be an absolute cell reference to the project manager input field in the settings worksheet. The other labels will be covered in the next paragraph. </p>

![Doc1 10](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/eb6d80ae-986c-4c30-96e7-2dac8fe6b0f1)

</br>
</br>

### 1.5. First configuration of the timeline

<p align="justify"> As the final step of the timeline setup, we also want to display the "Project timespan", which includes the "project_start" and end dates, as well as the "Project duration" using labels in the main header of the worksheet. For now, we are only going to enter an example date value for the "project_start" date and some placeholder values for the "project_end" date and duration. Nevertheless, later in this tutorial, we're going to transform this into a dynamically computed timespan. For the time being, we only need the "project_start" date as we need it to set up the timeline values, which means we will reference the start date as the initial date in the timeline. Make sure it has a clean date formatting, to do so go to the "Home" tab, "Number" section and open the dropdown list in the top. Here select the last option "More Number Formats..." to customize the date as you wish by selecting the category "Custom", however, the "dd-mmm-aa" customized format works fine. </p>
  
<p align="justify"> For the subsequent dates in the timeline, we make use of the workday function, as we want to entirely base its performance on actual workdays. The WORKDAY function returns a date, which is a specified number of work days after a given start date [ WORKDAY(start_date, days) ]. So for the start date, we simply reference the initial date in the timeline and as we just want to get the next workday in the cell right next to it, we enter 1. This will return a number because this cell has not the date formatting, copy over the date formatting and we can then fill all the subsequent cells in the timeline with the exact same formula, as we have used relative cell references. </p>

![Doc1 11](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/0b2c30c8-b6db-49fc-9523-c3b60dc8f8ee)

<p align="justify"> In the dark blue row below we want to display the initial weekday letter of the respective date. For that, we make use of the TEXT function which is perfectly suited to apply a specified text pattern to a date [ TEXT(value, format_text) ]. So let's reference the date in the cell above as "value" and specify the "format_text" as "ddd" which returns the initial three-weekday letters. As we are only interested in the first letter we're going to wrap this formula with the LEFT function to cut off the first letter from the left [ LEFT(TEXT(date_cell, "ddd"),1) ]. The 1 in the formula stands for the number of letters we want to keep starting from the left. Once this is done, just decrease the font size and add this formula to all the subsequent cells. This perfectly demonstrates that the timeline dates only cover the default work days, which are Monday to Friday. </p>

![Doc1 12](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/858bed5e-49e0-460f-89ed-a11df8a1c420)

</br>
</br>

### 1.6. Dynamic structuring of the timeline and chart area based on weeks

<p align="justify"> As the final step of this initial setup, we want to dynamically structure the timeline and chart area based on weeks so we'll need to create a visual separation between the weeks that is correctly displayed even if the timeline is shifted to the left or right. To do so, let's first create a dynamic name reference that dynamically references the date cells in the timeline; we select the initial timeline date cell, go to the "Formulas" tab and click on define name. The name of this reference will be simply "date", we set the scope to this worksheet only (so this name reference will only be valid in the scanned worksheet) and then we click into the reference down below which currently is an absolute cell reference with a dollar sign in front of both the column and the row. By removing the dollar sign in front of the column letter this name will now dynamically reference the date cell in whatever column it is called from, while row 10 is fixed. In a similar manner, we can also create a name that always references the next date in the timeline called "next_date". The reference you define in the formula bar below will always be relative to the cell that you had selected at the moment you opened the define name window; that means if we now, not only remove the dollar sign in front of the column letter but also change the column letter from X to Y (the next one) this name now always references the date cell in the next column, while the row 10 is still fixed. </p>

![Doc1 13](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/c280a41e-8e4a-4ee2-a2fc-1f02c1ea9d79)

<p align="justify"> We can now define a simple logical test. We check if the weekday of date is greater than the weekday of next_date [ (WEEKDAY(date,2) > WEEKDAY(next_date,2)) * 1 ]. With the simple workday focus setup that we implemented in our timeline this expression should always be false except for Fridays, so when multiplied by 1 it should be 1 for all the cells for which the date in the timeline is a Friday. Once we add that formula to all the cells in this row this proves to be true. So now we can set this formula apart and, instead of applying this formula to all the cells, make use of this expression to dynamically structure the timeline and Gantt area using a conditional formatting rule. 
  
<p align="justify"> We select the whole timeline and Gantt chart area, go to "Home" tab, hover over "Conditional formatting" and select "New rule". From the pop-up window select the last option "Use formula" to determine which cells to format and pass the formula that we have just created to identify all the cells for which we want to have a light gray right border to its right. That's all we need, we confirm and as you see not only does this beautifully add a weak base structuring to this area, but it also works no matter the start date. It would work even if you decide to add weekend days or holidays into your Gantt chart. </p>

![Doc1 14](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/bf7e08d7-8340-46b0-9e15-f53ff440233a)

</br>
</br>
</br>

## 2. Basic dynamic structuring of the "Item Details" section 

### 2.1. Visualizing the hierarchical structure of our items

<p align="justify"> Let's start by focusing on the basic details of our item setup. First, we will make a small separator row between the header and content and set the font size for the input area to 10. One of the most important input columns is the "Type" selection, which determines whether the item is a stage, task, or milestone. To allow the user to select from these options, we will create a drop-down list using the data validation window, select list and manually enter the values "S" for stage, "T" for task, and "M" for milestone. We will drag down the auto-fill handle to make this selection available in each row and set the text alignment to center. Next, we will enter some example items, such as a stage called "Planning" with some task items and a final milestone, and a second stage called "Implementation" with a similar structure. </p>

![Doc2 1](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/d02420cc-db3f-435f-97af-7b1690bee971)

<p align="justify"> We will create a dynamic name reference called "type" to reference the values in the type column for future calculations. We will set the scope to this worksheet only and remove the dollar sign in front of the row number to make the row dynamically changeable. This will make it easier to build and understand any formula that uses this value. To visually structure the Gantt worksheet, we will create a conditional formatting rule in the "Description" column that highlights every "Stage" row by setting the font style to bold and adding a mid-gray border at the top and bottom. To do so, in the input field of the new rule (select "Use formula" option) you must type in [ type = "S" ]. </p>

![Doc2 2](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/7f4ef159-7ebb-43db-b40c-91ff62ff5944)

<p align="justify"> We will also automatically indent the names of the tasks and milestones in the "Description" column to emphasize they are sub-items of the stage. For this, we will create an additional conditional formatting rule and use the custom category to add space characters for indentation. The rule will apply to the cells in which the "type" name that we defined earlier is either a task "T" or a milestone "M" [ OR(type="T", type="M") ]. To create the indented format head to the number tab, select custom and, here comes a cool trick; you can represent an indent value (a space) using the "@" symbol and add the desired number of spaces for indentation ["   @"]. Now we have a clear visualization of the hierarchical structure of our items. </p>

![Doc2 3](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/26fc09ce-cfb3-4863-9ca5-3923bafde23d)

</br>
</br>

### 2.2. "Color Indicator" column

<p align="justify"> To differentiate between stages and their sub-items, we will set up a default color indicator by adding two conditional formatting rules to the shrinked color indicator column right next to the "Description" column. One will fill a darker gray when the type equals "S," and the other will fill a lighter gray when the type is "T" or "M" (use the same formulated conditions as before). For an even cleaner formatting of this column let's remove the border formatting, so there are no grey lines clustering the column in its individual rows. </p>

![Doc2 4](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/0497e392-6d89-4aa4-80ea-578aef388a2d)

</br>
</br>

### 2.3. "Issues" column

<p align="justify"> In the last part of the basic item details setup, we will add a column that enables users to highlight issues. To do this, we'll make a simple drop-down list with a single symbol, aligned to the center. This ensures that users don't need to type anything manually. When there's an issue with an item, users can easily select the symbol, and when the issue is resolved, they can remove it. </p>

![Doc2 5](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/96118d8e-1db7-472d-99f4-79e7bcbbef92)

</br>
</br>
</br>

## 3. Role & Team Management system 

### 3.1. Creating the "Project Role Types" table

<p align="justify"> Let's set up a smart and scalable role and team management system that can manage up to eight project roles and an unlimited number of team members in the settings area. To do this, let's create a new section for the project role type management in the "Settings" worksheet, which will require five columns. We'll change the header to a colorful green and add gray border lines, just like we did for the "Project Details" section. Let's adjust the size of these columns to have a right and left padding and more space for the input columns. Since we want to create a list, we need column headers that can be copied from the "Project Details" labels. The first column will be for the "Role" name, and we prefer the text to be aligned to the middle and center. We can duplicate this using the autofill function and change the other two headers to "Code", which is the column to put in a shortcode that represents the role, and "Count" or "#", in which we will be automatically counting the number of team members assigned to this specific role. </p>

<p align="justify"> The role and code column are simple text input columns, so let's give them a light green background. As the count column will be a calculated column, we'll just set the background to a light gray. Let's add some example roles in here. The number of input rows is intentionally limited to eight to group your single team members into functional groups in order to remove complexity and get a better overview in case you have a large team. The second reason for this limitation, for a good visualization of different roles with different colors it is required to have a limited size of the color palette in order to keep them somehow distinguishable. The great advantage of this approach is that we can potentially add an unlimited number of team members to the whole project. </p>

![Doc3 1](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/3ed49abb-e739-4a68-ad4f-6ed076326828)

</br>
</br>

### 3.2. Creating the "Project Team" table

<p align="justify"> Once the roles table is created, copy it and paste leaving one column in between. Name this new table as "Project Team". The initial input column will be for the team member name, the second one for the role assignment, and then the third calculated column for automatically displaying the role shortcode. The "Name" column just requires simple text input, but for the assigned "Role" column, we want to grab all the defined roles from the previous table and make them available in a dropdown list. We want this dropdown selection to always display only exactly those values that are defined in the "Project Role Types" section but not the empty cell values. To create such a dynamic range reference that grows and shrinks based on the number of non-empty cells in a range, we can make use of the OFFSET function. </p>

![Doc3 2](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/4cdd8c52-926c-485b-8a85-1b54724efad3)

<p align="justify"> The OFFSET function always returns a range that starts at a starting reference, then we could add a row or column offset which we don't need in our case, but we need to define the height dynamically by counting the number of non-empty items in the maximum possible range and, for that, we can use the COUNTA function. This returns a dynamic array with exactly all the non-empty values and once you add another entry the array grows accordingly and when you remove it the array shrinks again [ OFFSET($H$9;;; COUNTA($H$9:$H$16)) ]. We can now use this formula directly to define a data validation drop-down list so let's open the data validation window, select list and we type the formula in the source field. </p>

![Doc3 3](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/c1a1e7ce-9f28-4747-8533-4853cdb4feb7)

<p align="justify"> For the code column we want to automatically display the respective shortcode based on the selected role; for such a task excel offers a wide range of LOOKUP functions but the go-to formula that I recommend is the good old INDEX MATCH functions' combination. The INDEX function requires you to select a range from which you want a value to be returned ("Role" column in the "Project Role Types" table) and then, to allocate a specific value in this previously selected range, the second required input is a row number. In order to get the correct row number we can use the MATCH function which simply looks up a lookup value, in our case, that is the assigned role ("Role" value of worker in "Project Team" table) in a lookup array ("Role" column in "Project Role Types"). Inside the MATCH formula the third element must be a 0, as we want an exact match for the function to return the correct row number for the INDEX function to correctly return the role code [ INDEX($I$9:$I$16; MATCH(O9; $H$9:$H$16; 0) ) ]. Now let's see what happens once we add this formula to the rows below. We get an NA error as no role is selected here in left cell. As this doesn't look too pretty, let's catch this error by wrapping the formula with the IFNA function, which allows us to return an alternative value (an empty string " ") in case this error occurs [ IFNA( INDEX($I$9:$I$16; MATCH(O9; $H$9:$H$16; 0) );" ") ]. </p>

![Doc3 4](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/4761c8fd-da45-4285-952e-61c2ca735d99)

</br>
</br>

### 3.3. Defining the "#" calculated column

<p align="justify"> Eventually, for the "#" column to automatically count the number of team members assigned to each role, we first check if the role cell is non-empty because otherwise, there's obviously nothing to count. We add this formula to all rows and instantly have an overview of the role distribution. Whenever we now add a team member and assign a role, this count is updated correctly. In case it is non-empty we use the COUNTIF function inside to count the number of values in that role range that equals the lookup role defined in the list. Otherwise, we simply return an empty text string (" ") [ IF(H9<>""; COUNTIF($O$9:$O$32; H9); "") ]. Let's add this formula to all rows and, like that, we instantly have an overview of the role distribution. </p>

![Doc3 5](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/2b1d132f-ee83-41a0-8a9f-3a7abd6322ae)

</br>
</br>

### 3.4. Using the smart lists created in the Settings workshet on the Gantt worksheet

<p align="justify"> Now let's jump to the Gantt worksheet, where we're going to make use of these newly defined smart lists. We now want to be able to select a team member via a dynamic dropdown list and then automatically display the role right next to it. First, select the "Team Member" column, let's add the team member drop-down list by applying the same exact technique explained a minute before. So, again, to create a dynamic range reference, go to data validation and enter the OFFSET function with the first cell acting as starting reference, being the first member in the "Name" column of the "Team Member" table in the "Settings" worksheet. We skip the next two arguments and then use the COUNTA function to count the number of non-empty cells in the whole team member name column [ OFFSET(Settings!$N$9;;; COUNTA(Settings!$N$9:$N$32)) ]. That's it, we can now assign a team member to each task and milestone. </p>

![Doc3 6](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/a05b50f5-a205-4673-8bcd-712b0965839a)

<p align="justify"> In order to automatically display their respective role in the left column, we only need to look up the selected name in the project team member list and return the role shortcode. For this let's first set up a dynamic name reference for the team member just as we did before for the "type" name reference. Call it "team_member", the scope is the Gantt worksheet and to make it dynamic we are going to remove the dollar sign in front of the row number (so the reference is fixed in the column but variable between its cells / rows). In the "Role" column of the "Item Details" section we define again a new INDEX MATCH fuction wrapped in an IFNA fuction to look up the respective row. The return array is the shortcode column of the "Project Team" table in the "Settings" worksheet, so we apply the INDEX function to this column, and then to get the correct row we apply the MATCH function to look up the team member in the "Name" column as an exact match (0). In case this returns an N/A error let's just return an empty string again (" ") [ IFNA( INDEX(Settings!$P$9:$P$32; MATCH(team_member; Settings!$N$9:$N$32;0) ); "") ]. Once we add this formula to all the rows, decrease the font size, change the color to a gray tone and also add some indent. We can even increase the size of this column a little bit and now this is an almost perfect implementation of this team member and role assignment. </p>

![Doc3 7](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/188563ba-42bc-4883-afb4-28f59df4a106)

<p align="justify"> There's only one last thing that you could add to this feature to make it fully complete. Jump back to the "Settings" worksheet and notice how we need to delete one column to make the "Gantt" button come back to its original place, as we have added one additional column in the process. Now the cherry on top you may want to add is a user-friendly way of removing team members from the "Project Team" table which might happen from time to time. Removing a team member in this table is no problem but, what if the team member you deleted  is currently assigned to an item in the Gantt worksheet, then the role is empty because the lookup results is an N/A error. However, the name is still written in the cell. So it might be helpful to visually indicate that the assigned team member is no longer part of the team and the item has to be reassigned. My idea here is to add a conditional formatting rule that checks if a MATCH function lookup of that name in the name list returns an N/A error using the ISNA function and, if that is the case, we change the font color to a warning red [ ISNA( MATCH(team_member; Settings!$N$9:$N$32; 0) ) ]. </p>

![Doc3 8](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/a7cb92c8-f21f-4bb7-9622-ca9fa054afb8)

</br>
</br>
</br>

## 4. Manual Planning with independent Start or End Dates

### 4.1. Manual plannig mode given an Independent Start Date

<p align="justify"> Now let's move on to implementing the basic manual planning with independent start or end dates. The most crucial columns for the actual plan will be the two calculated columns for the start and end dates, as all the date-related visualization in the chart area will be based on these two columns. The reason why the start and end date need to always be calculated is that we're going to have different input options. As previously mentioned, the calculation of the start and end date will always be based on two inputs: one input is the number of required workdays, and the second input is either an independent start date, an independent end date, or a dynamic start or end date that is dependent on another item. Of these two calculated start and end date columns, one will always be defined directly from the date section, and then the other one will be calculated based on that and the defined number of workdays. Before we start setting this up, let's make sure that all the date columns, including the base plan columns, have the same clean date format (dd-mmm-aa). </p>

![Doc4 1](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/a0ad9cb7-5a0d-4810-9b01-e4018ee92583)

<p align="justify"> Let's start by setting up the simplest of all options, which is providing an independent start date and the number of required workdays. In order to make these cell values easy to reference in other formulas, we define dynamic name references for the independent start date (you can call it "ind.start"). Always keep in mind to remove the dollar sign in front of the row number. Then we do the same for the workdays, but this time we not only make it dynamic, but we also want to make sure that this name reference will always have a default value of 1 in case the cell is empty. We can easily do that by returning the maximum of the cell value and 1 [ MAX(Gantt!$R15; 1) ]. This has two big advantages. First of all, it is guaranteed that an item will become instantly visible once either a start or end date is provided, even before the number of workdays is entered. Second, we don't need to actively enter a workday number for milestones. </p>

![Doc4 2](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/2f0ed4e1-625f-4804-b935-33c29a5f3ba4)

<p align="justify"> As we will need to reference the calculated "Plan start" and "Plan end" for the Gantt chart visualization, let's also add dynamic name references for these. I explicitly call them "plan_start" and "plan_end", as these will be the ones that will always represent the actual plan in one constant place. The formula for the "Plan start" in this scenario is pretty simple. We first check if the independent start date is non-empty [ IF(ind.start<>""; ], and if that is the case, we return the independent start date; Otherwise, an empty string [ IF(ind.start<>""; ind.start; "") ]. So this is the date directly taken from the input section, whereas the "Plan end" date has to be calculated based on the "Plan start" and the workdays. In a similar manner, we also check if ind.start is also non-empty because only then the "Plan start" is also not empty [ IF(ind.start<>""; ]. Then. we're going to make use of the WORKDAY function, put in the "plan_start" as start date and add the "workdays" reference minus one. The reason why we have reduced this number by one is that the planned start date is already one of these 10 workdays, so the "Plan end" date must be one day less [ IF(ind.start<>""; WORKDAY(plan_start; workdays-1); "") ]. The alternative value in case ind.start is empty is again an empty string. </p>
  
![Doc4 3](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/7c0f04ef-0ac1-4c23-90b7-d5384e4127ba)

<p align="justify"> Let's do a quick check if the calculation is correct by looking at the timeline. As you see, the 10th to the 21st of May are exactly a 10 workday time span, so it works correctly. When we take a closer look at this formula, there are two crucial things we always have to keep in mind. First, every time we calculate one of these plan dates by adding or subtracting the workdays from the other date, it is crucial that we reference the other calculated plan date and not the input date like ind.start; because otherwise, we will end up creating circular references at some point. The second aspect to consider is, as both these formulas will get a lot more complex over time and also we are going to reuse certain calculations, it is advisable to transform all sub-calculations into what I call named calculations. What that means is that instead of entering these sub-calculations directly into the formula itself, we gonna create names for them in the name manager and only have to reference the name in the bigger formula for the same exact result. That not only makes the bigger formulas easier to read and is reusable, but also allows us to manage and potentially adjust each sub-calculation at one central place in the name manager. </p>

<p align="justify"> Now let's use this to manually set up the schedule for the tasks and milestone of this stage. As you see, as soon as we have entered the independent start date, the "Plan start" and "Plan end" are instantly displayed as the workdays value is set to 1 by default. And for the milestone, we only need to enter the independent start date and can leave the workday cell empty because we want the milestone to start and end on one particular date. With this example schedule set up, we can now create the default visualization in the chart area. </p>

![Doc4 4](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/3b872fb4-f2ba-4e42-aee6-ae750368dad7)

</br>
</br>

### 4.2. Primal Gantt planning visualization

<p align="justify"> Right below the timeline, we have all the information prepared to develop a formula that lets us determine whether an item is within its plan in a given cell or not. Because we now know what the "Plan start" is, what the "Plan end" is, and what a respective date in the timeline is. These name references work correctly from every cell in the chart area. To figure out if a cell should display something or not, all we need to do is check if the date is within the planned time span. So, if the "date" is greater equal to "plan_start" and smaller equal to "plan_end", this formula returns either true or false [ AND(date>=plan_start; date<=plan_end) * 1 ]. And for demonstration purposes, you can display either a 1 or a 0 by multiplying the formula by 1. When we add this to all rows we can verify it works perfectly for all defined items. For this formula again, it makes sense to transform it into a named calculation. So let's copy it and create a new name that will be called "item_in_plan". Set the scope to this worksheet and paste the formula. Now we can replace this formula and have one short expression that works in every cell of this Gantt chart area, and that is all we need to implement the default visualization. </p>

![Doc4 4](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/4511195a-6e98-4132-8060-b9661846df89)

<p align="justify"> One big goal of this template is to have the stages, tasks, and milestones all differently visualized. While the stages and tasks can be perfectly visualized using a cell color fill without written cell content, we want the milestone to display an actual milestone symbol. For that purpose, I have prepared two Unicode icons, one for an open and one for a completed milestone (◆ - ⚑). For now, let's copy over the open milestone unicode icon (◆) and start building the milestone visualization as an in-cell formula. In the formula, we only have to check two conditions that have to be true. First, the item "type" has to be a milestone ("M"), and second, the item has to be in plan in this cell ("item_in_plan" = 1). If that's the case, we just print the milestone symbol (◆). Otherwise, an empty text string [ IF( AND(type="M"; item_in_plan); "◆"; "") ]. Let's add this formula to the full chart range. In addition, we're going to align the cell content to the middle and center and slightly increase the font size. </p>

![Doc4 5](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/031bd157-0de4-4686-9950-3d3525127d0d)

<p align="justify"> For the stages and tasks, on the other hand, we want to create Gantt bars by filling the respective cell backgrounds with some color. For that, we can directly jump into the conditional formatting rule manager and create two new rules. Let's create a default rule for the stage, just like for the milestones, this formula checks if the "type" is a stage ("S") and if the item is in plan in the respective cell ("item_in_plan" = 1). If both these conditions are true, we want to add this darker gray background fill color to the cell. Let's apply this rule and move it down in the rule order, which simply means that this rule is now applied after the two other rules. We duplicate this rule to create a similar rule with a lighter gray for the tasks [ AND("type"="T"; item_in_plan) ]. Once we apply this rule, both the stages and tasks are now beautifully visualized. The only thing left now is to color the milestone symbol with the stage given color. So let's jump back to the conditional formatting rules manager, duplicate one rule, change the "type" condition to milestone ("M"), clear the cell fill formatting and change the font color to a darker gray. </p>

![Doc4 6](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/bb425202-df7a-4df5-a8c0-d83d8f6be9ed)

<p align="justify"> Before we close this conditional rules formatting manager, I need to show you two concepts in here that are crucial in terms of worksheet efficiency. At first, always add a check in this "stop if true" checkbox if a rule and all the following rules are mutually exclusive. What that means is, in this case, if the stage coloring rule here is applied because its conditions are true, then the two following rules don't have to be checked because they cannot be true at the same time. That will save you a lot of computational resources, especially with a larger number of rules in here. And the second concept somehow builds upon this "stop if true" feature and is even more crucial. When we take a look at the default coloring rules previously defines, we see that each of these does the same "item_in_plan" check, which is super inefficient. To avoid this, we can add one additional rule that does this "item_in_plan' check once at the beginning. And only if the cell passes this test, we continue testing the following rules. This is called a "stopper rule" because what we do in this new rule is checking the opposite of the condition that we want to be true. So in this case, we check if there's no "item_in_plan" [ NOT(item_in_plan) ]. Place it right on top of the coloring rules, activate the "stop if true" checkbox and erase the "intem_in_plan" condition check from the coloring rules. Now this whole setup is super efficient, and we are perfectly prepared for scaling this with some additional coloring rules later. </p>

![Doc4 7](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/68a7863e-be99-4138-bb96-3879ac0ccf58)

</br>
</br>

### 4.3. Manual planning mode given an Independent End Date

<p align="justify"> Now that we have set up the default visualization, let's take a look at the other manual scheduling option, which is instead of entering an independent start date and doing a forward scheduling, entering an independent end date and doing a backward scheduling. Let's reproduce the same exact time span for this stage by entering the end date that is calculated in the "Plan end" cell at the moment, which is the 21st of May. At the moment, the "plan_start" and "plan_end" formulas are only able to work with the independent start date input. </p>

<p align="justify"> First, we're going to create a dynamic name reference for the independent end date ("ind.end"), remove the dollar sign in front of the row number, and then jump straight into "plan_start" and "plan_end" formulas. This time, we start by modifying the "plan_end" formula. At first, instead of returning an empty string if ind.start is empty, we now continue with a second nested IF statement that checks if "ind.end" is non-empty because, if that is the case, we want to directly get the value of "ind.end" and only otherwise we're going to return an empty text string (*). </p>

```
(*) = IF(ind.start<>""; plan_end_calculation;
                        IF(ind.end<>""; ind.end; ""))
```
![Doc4 8](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/52065492-395f-4e30-a27e-a0966e3d980d)

<p align="justify"> Let's also update the formula in the other rows of that column and, this time, the "plan_start" has to be calculated backwards based on the "plan_end". First, we check if "ind.end" is non-empty and, if that's the case, we use the WORKDAY function, put in the "plan_end" calculation as the start date of the WORKDAY function and then add the negative number of "workdays" (-1, the idea is again going day by day but backwards); so we multiply the workdays by -1.  Since the "end_date" is already one of these 10 "workdays" we have to add one again to make it only go back 9 workdays, otherwise, if "ind.end" is empty we just return an empty text string (*). We now have successfully reproduced the exact same time span that we had before so, again, let's take that sub calculation and create a new name reference that we call "plan_start_calculation" (**). Update the other rows in the "Plan start" column and test if it recreates these time spans by providing only independent end dates. </p>

```
(*) = IF(ind.start<>""; plan_end_calculation;
                        IF(ind.end<>""; WORKDAY(plan_end; (-1)*workdays+1); ""))

(**) plan_start_calculation = WORKDAY(plan_end; (-1)*workdays+1)
```
![Doc4 9](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/b48d64f8-8dd4-4e78-b135-a947f0e02809)

<p align="justify"> We have intentionally implemented both the "plan_end" and "plan_start" formulas using a hierarchical logic, which means these formulas are written in a way that the "ind.start" value will always override the "ind.end" value which will be ignored. To make the manual planning mode more intuitive we're going to add a conditional formatting rule to the "ind.end" column that simply checks if "ind.start" is non empty [ ind.start <> "" ], if that's true the "ind.end" value will be inactive and we indicate that by giving it a mid-gray font color. Now we instantly know which one of both independent dates are we currently using. Let's overwrite the stage and task items with the original "ind.start" date, leave the milestone defined via the "ind.end" date and we quickly set up a second stage. The first task using a forward scheduling and the other two tasks being backward scheduled. That way we can perfectly see how the bar now grows backwards as soon as we set the number of workdays to a number that is bigger than 1 same for the next task. 
  
![Doc4 10](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/d77a44dc-31fd-42af-870d-87af5a0fb855)

</br>
</br>

### 4.4. Introduction to the Dependency Engine for Dynamic Planning

<p align="justify"> It's the perfect time to explain how quickly we can automate stage calculations and visualization based on its items. As I have no intention to override the calculated "plan_start" and "plan_end" formulas. The clean approach here is to work with some simple formulas in the input section; for that, I have prepared two standardized placeholder formulas that are easy to use. All you need to do is copy the "Auto stage Ind.Start" placeholder formula into the "ind.start" cell of a stage, and then replace this placeholder with the planned start range of the items in that stage (*). In the same manner we can dynamically calculate the number of workdays using the given placeholder formula ("Auto Stage Workdays") which references the "Plan start" and then calculates the net workdays between this "Plan start" and the maximum of the items, within that stage, "plan_end" dates (**). We are now free to change any of the tasks or milestone schedules and the stage timespan will always be automatically updated. </p>

```
(*) = MIN(ITEMS_START)
(**) = NETWORKDAYS(plan_start; MAX(ITEMS_END))
```

![Doc4 11](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/ed6b123a-49f7-4a47-8e0a-39cd475729c7)

<p align="justify"> I think the usage of formulas in these input cells somehow makes it necessary to be highlighted. Let's add a new conditional formatting rule to the "ind start" column that checks if the cell contains a formula, let's make the cell reference fully relative and the formatting will be a strong blue font color (as always, formatting, is up to you). Let's also include the workdays column into this rule and that makes these formula inputs clearly distinguishable from regular inputs [ ISFORMULA(cell without any $) ]. Once we have built such an auto-stage formula in one stage we can simply copy it over to another stage since the range references are relative. In case you have a different number of items all you need to do is adjust this referenced range with your mouse and you are good to go. </p>

![Doc4 12](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/65f3e766-f764-4623-b39e-bd95c9347608)

<p align="justify"> What really amazes me about this general setup is that we can now collapse the date input columns and have the workdays as key information still visible. Whenever we want to adjust the number of required workdays for a task there is no need to manually update the stage workdays as they are updated automatically with this auto-stage workday's calculation. This formula is not simply adding up all the workdays of stage items, but actually computing the maximum time span of all items; by having this function it is fully able to handle whatever items overlap with the time span (*). For the auto-stage "ind.start" value we can make it even more obvious that this is not a real "ind.start" date by adding an additional rule that only focuses on the "Ind start" column and is also limited to Stage rows. Only when the cell is in a Stage row and it is a formula, instead of displaying the date itself, a really cool trick is to just override it with a text value which will be "Auto" [ AND(type="S"; ISFORMULA(ind.start)) ]. I think this is a beautiful and intuitive setup, a first take on what's about to come.</p>

```
(*) = NETWORKDAYS(plan_start; MAX(ITEMS_END))
```

![Doc4 13](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/18577ae7-2b89-40a2-afea-a9c803458a60)

</br>
</br>
</br>

## 5. Dynamic Planning with Dependency Engine

<p align="justify"> This manual planning with independent start or end dates is pretty helpful but it is nothing compared to what we're going to build. We are going to add the core feature of this template, an intuitive and powerful dependency engine that allows you to make your project plan fully dynamic. We are also going to implement all four possible types of dependency connections between items. The section that will allow you to create all kinds of dependency connections between items quickly is the "Dependency" part within the "Actual plan" section. We are going to build up the whole dependency logic using the second task P2 in first stage as example. Instead of relying on one of the manually defined independent dates, we will make that task P2 dependent on task P1. So let's clear all the inputs for the subsequent items, remove the independent end dates and since we want to make the task P2 dependent on the task P1, we're also going to remove the independent start date for this one. The workdays can just stay in there as we're going to use them in the same way as we did before. Let's also temporarily move the timeline one week to the right so that not only the forward-looking dependencies are visualized, but also the backward-looking dependencies. </p>

</br>
</br>

### 5.1. Elementary construction of the ID's System

<p align="justify"> The key idea for making an item dependent on another item is linking it to the other item's id. This obviously requires us to create a reliable id system in advance, and for that, we start by defining a dynamic name reference for the cells in the "id" column within the "ID's" section. We call it "id", set a scope, and as usual, remove the dollar sign in front of the row number. As we don't want to manually enter any ids, we somehow need them to be automatically created and self-sustaining in case we insert or delete a certain row. That means we want to have an id system in which each id dynamically represents the position of its item's row in the whole list of items as a unique number. The straightforward solution of simply entering and autofilling a sequence of numbers directly is not sufficient in that case because as soon as we insert a row and use the fill down option, we get duplicates. Setting the first id to 1 and then for all subsequent ids just referencing the previous id and incrementing it by 1 is also not fully working as the id below the inserted row unfortunately keeps referencing the same id. So we have to create something that is way more robust and self-sustaining. </p>

<p align="justify"> The key to achieving that is to give each cell all the information about the previous values in its column. We can easily create a named range reference that we call "prev_col_range", while having selected any cell in the column, define the range starting from the header cell (included) and ending at the cell right above of the selected one. At the moment, this is a fixed range, but by removing the dollar sign in front of the last cell's row is going to make this a range that will dynamically grow and shrink depending on the cell from which it is referenced. Furthermore, as we might make use of this in other columns as well, let's also remove the dollar signs in front of the column's letter. </p>

<p align="justify"> The reason why we included the header in the range is because that ensures that even if we insert a row on top of the initial row it will become part of that range and not be excluded. Now we can use the name reference to enter the formula that will dynamically create the correct id for each row. In this formula we first need to check if the current row is the initial row in the list of items, that's the case whenever the row of the cell itself equals the row of the header plus two (two because we have the separator row in between so the initial row will always be two cells below the header cell). If that is the case the id has to be one and, for all the subsequent cells for which that is not the case, we simply compute the maximum of the previous column range and increment it by one, simple as that (*). Let's drag the autofill handle down and that looks good so far, but the actual power of this becomes visible once we insert or delete a row and fill the values down you see all the id values get perfectly updated. </p>

```
(*) = IF(ROW(id)=ROW($C$14)+2; 1; MAX(prev_col_range)+1)
```
![Doc5 1](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/f44cc1e5-d7c9-4dec-9c26-4fe0dba62185)

</br>
</br>

### 5.2. Crafting the Dependency Information Display

<p align="justify"> Now we can jump over to the dependencies section and make task P2 link to the id of task P1, this link should always be created by actually referencing the respective id cell in order to make sure that the linked id number is correctly updated in case the other item's id changes due to an inserted or deleted row. To help the user instantly see if the linked id is an actual cell reference and not just a fixed number typed in, we make use of the previously defined conditional formatting rule that gives the cell content a blue font color in case it contains a formula. We increased the conditional formatting rule range by adding a comma and selecting the "d.id" column, and now this reference id is colored blue which tells us it links to another cell. </p> 

<p align="justify"> Now that we have set up a way to create a robust link to another item's id, let's make this cell value dynamically referenceable by creating a name reference called "d.id" (d stands for dependent). The first thing we can do with this dynamic reference id is to extract some fundamental information about the other item. The first information we are interested in is the position of the other item, so whether it is positioned above or below the dependent item's row. To indicate that visually we're going to make use of the two unicode arrow icons. The formula for this is quite straightforward, at first, we check if "d.id" id is non-empty and if that is the case we only have to check if "d.id" is smaller than this item's id which means it's positioned above, we thus insert a placeholder for arrow up and alternatively, in case "d.id" is greater than the depended item's id the other item is obviously positioned below, so here we want to display the arrow down. As a fallback option, we just return an empty text string (*). </p>

```
(*) = IF(d.id<>""; IF(d.id<id;"↖"; IF(d.id>id;"↙";""));"")
```
![Doc5 2](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/12f1a79a-cff2-4edf-a329-3f22e33d21a3)

<p align="justify"> In the next column we want to automatically display the type and name of the linked item. This is a classic look up problem so we can use the MATCH function to look up the "d.id" value in the "id" column to get the correct row number, and then use the INDEX function to return the respective value from the "Type" and "Description" columns. As we're going to need the row number of the other item at multiple occasions let's just start with the MATCH part and make a named calculation for the lookup value; we pass "d.id" as the desired value and then, for the lookup array, we can pass the whole column B ("Id" column) as an absolute reference with dollar signs. We also add a 0 to have an exact match (*). This returns 16 in this example which indeed is the row number of the linked item, so let's copy that formula and transform it into a named calculation that we call "d.row". For extracting the other item's type we use the INDEX function to take a look at the column E ("Type" column) and return the value from the row calculated in "d.row" (**). We can save this new INDEX calculation as "d.type". Lastly, the linked item's name it's basically the same approach let's call it "d.name", paste the INDEX formula we just used but replace column E ("Type") with column F ("Description")(***). </p>

```
(*) "d.row" = MATCH(Gantt!d.id; Gantt!$B:$B; 0)
(**) "d.type" = INDEX(Gantt!$E:$E; Gantt!d.row)
(***) "d.name" = INDEX(Gantt!$F:$F; Gantt!d.row)
```
![Doc5 3](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/f1b4830a-2bef-40a6-a918-71e829a26b84)

<p align="justify"> Now we have all this information about the other item available as named calculations. Let's also decrease the font size a bit and change the font color of all this information to a dark gray to show these are secondary lookup information. We are now prepared to create a concatenated string that contains "d.type" in brackets, followed by "d.name". This string correctly displays the information but once we add this formula to some row that has not a "d.id" reference this will result in an N/A error. In order to catch this and potential other errors, let's wrap this expression in an IFERROR function that displays an empty string in case an error occurs. An additional aspect to consider is that the other item's name could potentially be way longer than the space left in the cell. To avoid this problem adjust this formula; instead of just directly printing the full name we first are going to check if the length of this name exceeds a certain threshold (for example 7 characters) and, in case it exceeds that threshold we're going to apply the left function to extract the first threshold number of characters, so in this case 7 followed by some dots (*). Let's do a final test for the whole "d.item' output by changing the referenced id once again, and the information displayed matches the information for the new dependent item. </p>

```
(*) = IFERROR("("&@d.type&")"&@ IF(LEN(@d.name)>7; LEFT(@d.name;7)&"…"; d.name);"")
```
![Doc5 4](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/6cc74eea-734d-4e5d-a43a-c9b235e1203a)

</br>
</br>

### 5.3. Building the Dependency Connection ENgine

<p align="justify"> Let's move on to the heart of the dependency engine which is the selection of the type of dependency connection we want to have in the column "d.conn". Head to the "Data" tab and add a drop-down list using "Data Validation" which will let us select one of four different connection types: finish to start, start to start start to finish, and finish to finish. The first letter always refers to the other item and the second letter to the current item. A great way to provide some basic information and instructions for an input column like this, is to select the header cell, then open the data validation window and go to the "Input message" tab. This section allows us to enter an informative message that is displayed whenever this cell is selected. The column next to "d.conn" is "d.lag" and will allow us to define a lag by which we want the lagged item to be shifted to the right in case we have a positive value or to the left if it is negative. Let's set the text alignment to center for both these columns, adjust the column size and create dynamic name references for both these columns. Call the first one "d.conn", the second one "d.lag" and make them both row relative. </p>

![Doc5 4](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/77f20079-ea41-4aaa-b1b4-99a77ecddb41)

<p align="justify"> Now we have all the values, variables and information available to make the dependency logic part of the "Plan start" and "Plan end" columns' formulas. If you remember what I already said, the hierarchical logic works as it follows; both the "Plan start" and "Plan end" formulas always start by look at the "ind.start" value and only if that cell is empty they consider the "ind.end' value, and only if that cell is also empty they will now consider the dependency section and do the date calculations based on what is in there. </p>

```
(*) Plan start column formula (before) = IF(ind.start<>""; plan_end_calculation;
                                                          IF(ind.end<>""; plan_start_calculation; ""))
```

<p align="justify"> So let's start with the "Plan start" formula. We are going to replace the final empty string with a completely new expression. In this new formula portion, first we will check if "d.id" is non-empty, otherwise, there is nothing that could be calculated. In case it is non-empty we can take a look at what connection type has been selected for the "Plan start" formula (d.conn = "FS"). The most common one is the finish-to-start dependency, in that case you want to do the respective calculation for which we simply enter a placeholder for now ("fs_calc"). Alternatively, if the selected dependency connection type is start to start we want to do the respective start-to-start calculation ("ss_calc"). In case it is one of the other two dependency connection types, then we know that the "Plan end" date will be tied to the other item start and we need to do the backwards calculation based on the "Plan end" date and the number of required workdays. Guess what? we have already define this backwards calculation as "plan_start_calculation". Now we only have to close these brackets for the last two if statements, then provide a fallback value in case the id is empty and eventually close that bracket as well. For now with finish to start or start to start selected (FS or SS) the respective placeholder appears correctly and, once we select one of the other two types we get an error because the formula assumes that the "Plan end" date is already available, which is not the case. Yet to catch these errors, let's wrap this formula inside an IFERROR statement just to make sure it will always display an empty string in such case (*). </p>

```
(*) = IFERROR( @IF(ind_start>0; ind_start;
                      IF(ind_end<>""; plan_start_calculation;
                                      IF(d.id<>"";
                                                IF(d.conn="FS"; "fs_calc";
                                                IF(d.conn="SS"; "ss_calc";
                                                plan_start_calculation));"")));"")

```

![Doc5 6](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/1968bedf-83a1-4bfe-aa6b-c9b93258caab)

<p align="justify"> For the "Plan end" formula, we're going to implement the same exact pattern. Instead of returning an empty string we first check if "d.id" is non-empty and if that is the case we now check if "d.conn" equals one of the two finish dependency connection types which are start-to-finish (SF) and alternatively finish to finish (FF), otherwise, in case it is one of the other two values we need to calculate the "Plan end" based on the "Plan start" and the number of required workdays, which is already covered in the "plan_end_calculation". Let's close both these if statements, enter a fallback value in case "d.id" is empty and we can also directly wrap the whole formula in an IFERROR function to catch potential errors. The basic logic of when to do which calculation is now implemented when switching between different "d.conn" values, the placeholders perfectly show us which one, the "Plan start" or "Plan end", is the one being calculated based on the other item (*). </p>

```
(*) = IFERROR( @IF(ind_start<>""; plan_end_calculation;
                                  IF(ind_end<>""; ind_end;
                                                         IF(d.id<>"";
                                                                    IF(d.conn="SF"; "sf_calc";
                                                                    IF(d.conn="FF"; "ff_calc";
                                                                    plan_end_calculation));"")));"")
```

![Doc5 7](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/484ea1ab-0f51-4528-b968-a7d9ed4ec727)

</br>
</br>

<p align="justify"> With all the preparation ready we are set to replace these placeholder calculations with some actual calculations. For the finish-to-start calculation (FS) we need to set the "Plan start" date to the workday after the "Plan end" date of the connected item. The WORKDAY function is the one that allows us to do exactly this calculation, we only need to grab the "Plan end" date of the connected item value but, as we already know the row of the other item, we can simply return this end date value using the INDEX function by entering column T as the return array and "d.row" as the row. With this, we now have a start date defined in the WORKDAY function and we only need to specify the number of workdays we want to add to this date (1) plus the lag days as "d.lag" (*). The most important aspect is that when we change the duration of the first task, the plan dates of the second task are updated correctly. </p>

<p align="justify"> For the start-to-start (SS) calculation we can now simply copy over that formula as we only need to make some tiny adjustments. For this one we want both items to start together so this time, instead of grabbing the other item's "Plan end" date, we now need to grab the other item's planned start date from column S and for the day's argument, we're going to remove the one because is not the next day but the same day that we want (it will only be shifted by a potentially given lag) (**). For better maintainability and less complexity in this huge formula, let's transform both these expressions into named calculations. The first one will be called "fs_calc" and the second one "ss_calc". </p>

```
(*) "fs_calc" = WORKDAY( INDEX($T:$T; d.row); 1+d.lag)
(**) "ss_calc" = WORKDAY( INDEX($S:$S; d.row); d.lag)
```
![Doc5 8](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/ea4ec9c2-16f4-4e20-8dee-2b2a0a4a49ea)

<p align="justify"> Let's jump right over to the "Plan end" formula. We're now going to create similar expressions for the two finish dependency connection types. For start-to-finish calculation (SF) we want to have the "Plan end" date of the item one day before the planned start of the other item, plus a potential lag. That means we're going to make use of the WORKDAY function again and, as the date to start with, we're going to grab the other item "Plan start" by applying the INDEX function to column S with "d.row" as the returned row. Since we want to calculate the workday right before that date we add -1 plus the potential lag value (*). Let's close that statement, hit enter, and change the dependency connection to start to finish to see if it works. </p>

<p align="justify"> Finally, for the finish-to-finish dependency connection type (FF) let's copy over the "SF" expression, remove the -1 and grab the other item's date from column T instead of column S, as we need the "Plan end" date. Like that, both items now will finish together on the same day, but potentially shifted by a given lag (**). We can now transform both these new expressions in the "Plan end" formula into two new named calculations, the first one called "sf_calc" and the other one "ff_calc". </p>

```
(*) "sf_calc" = WORKDAY( INDEX($S:$S; d.row); (-1)+d.lag)
(**) "ff_calc" = WORKDAY( INDEX($T:$T; d.row); d.lag)
```

![Doc5 9](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/d75fd4da-762c-4c78-a939-47ae7fd0688b)

</br>
</br>

<p align="justify"> In this whole dependency logic there is only one thing left to consider. So far we haven't considered that the user might not select a dependency connection, so the cell is empty. With the setup we've built that will result in a circular reference, that's the reason why a warning pops up. This occurs because both the "Plan start" and "Plan end" formulas try to calculate their dates based on each other, as neither of them is able to find one of their two relevant dependency connection types in there, so the "plan_start_calculation" and the "plan_end_calculation" are executed at the same time. We can easily avoid this by setting a default value in case "d.conn" is empty, in my opinion, finish-to-start is the most intuitive choice here. The quickest approach is making "FS" the default value by applying the finish-to-start calculation ("fs_cal") within the "Plan start" formula either if "d.conn" equals "FS" or "d.conn" equals nothing (" "). </p>

![Doc5 10](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/930e9fea-17a1-48c3-9cfa-2e6fb87fefde)

<p align="justify"> Update all the cells in the "Plan start" and "Plan end" columns with the newly built formulas. Do the same for the "d.item" column and, since finish-to-start is the default value in the dependency calculation, I suggest just pasting it as an actual value for all cells in the "d.conn" column and then making the columns behind the "d.id" column invisible. As long as no "d.id" is referenced we can achieve that with one simple conditional formatting rule; this rule will check if "d.id" is empty, in that case, change the font color to white, that way nothing is visible. </p>

![Doc5 11](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/51ef1100-3fb7-4092-8129-551a8ddbd720)

<p align="justify"> Let's recapitulate, the input logical hierarchy we've implemented during this whole formula development. We have built them into both the "Plan start" and "Plan end" columns and, with these updated versions of both formulas, the dependency functionality or mode has the last place in this hierarchy. This means it is overwritten by both the "ind_end" date and "ind_start" date, which are manually entered by the user. Finally, in the last layer, the "ind_end" date is overwritten by the "ind_start" date. For the dependency connection this means, it becomes inactive whenever at least one of the independent dates is non-empty. So let's add the same graying out conditional formatting to the dependency section as we did with the independent end date column ("Ind.End"). This rule checks if at least one of the independent dates ("ind_start" or "ind_end") is non-empty, in case that is true we set the font color to a mid-gray. When the rule is applied we can see it makes the "d.conn" values appear again, so let's just move it below this rule. The final formulas in the "plan start" column (*) and "Plan end" column (**) are written below. </p>

```
(*) Plan start column (final result) = IFERROR(@IF(ind_start<>"";ind_start;
                                                  IF(ind_end<>"";plan_start_calculation;
                                                                IF(d.id<>"";
                                                                            IF(OR(d.conn="FS";d.conn="");fs_calc;
                                                                            IF(d.conn="SS";ss_calc;
                                                                            plan_start_calculation));"")));"")
```
```
(**) Plan end column (final result) = IFERROR( @IF(ind_start<>""; plan_end_calculation;
                                                  IF(ind_end<>""; ind_end;
                                                                          IF(d.id<>"";
                                                                                      IF(d.conn="SF";sf_calc;
                                                                                      IF(d.conn="FF";ff_calc;
                                                                                      plan_end_calculation));"")));"")
```

<!--
### 5.4. Dependency planning. Example

Great let's do some example planning with the dependency engine for all the items that we have set up so far we gonna make this task dependent on the preceding task with a simple finish to start dependency and see how the task has one workday by default and as soon as we enter a number higher than that it grows to the right in a similar manner we could now make the milestone dependent on that last task in that stage by referencing id4 but in cases like that where we simply want to have consecutive item dependencies there's a beautiful shortcut you can use which is the autofill function simply select the dot id of the preceding item and then drag this auto fill handle down and since this is a relative reference in there this auto filled reference now links to the next id which is 4. for milestones i generally prefer a finish to finish dependency as a milestone is often logically tied to the successful completion of one or multiple tasks for the second stage we can start by copying over the autostage formulas because we already know the amount of items in that stage in case that stage had a different number of items we could easily just click into the formula and adjust the referenced range to cover all of the items and the same for the auto stage workday formula using this stage i want to demonstrate how you can easily make a stage dependent on another stage for example with a finish to start dependency connection all you have to do is to determine which one is the initial item in the second stage then link this item to the first stage and make the subsequent items directly or indirectly dependent on the initial item for consecutive dependencies within the stage we only have to link the second item to the first one and then we can use the autofill function to create consecutive dependencies for the last two items let's change the milestones dependency connection type to finish to finish and enter the number of required workdays for the tasks see how easy that was both these stages are now tied together and every change that impacts the first stages plan and date like for example an increased duration or change in dependency structure within that stage will make the second stage automatically adjust its schedule accordingly now that was a classic forward scheduling dependency connection between stages because we have made the start of the second stage dependent on the other stage we can also create a backward scheduling dependency for these stages by making the plan and date of the second stage dependent on the other stage for example as a finish to finish dependency that would mean the stage and date is fixed and every increase in duration of this stage would make the stage grow backwards to achieve that we have to reverse the dependency logic in that second stage initially we determine a final item in that stage in this case this would most likely be the milestone then link this final item to the other stage and make all the other items of that stage directly or indirectly dependent on that one so we update these dependencies accordingly by clicking into the did cell of the milestone and dragging it to the id of the first stage now this is connected to the first stage with the finish to finish dependency then we make that last task dependent on that milestone with the finish to finish dependency and for the upper two tasks we can use the autofill function again and only have to update the dependency connection type to start to finish for both of them whenever we increase the duration of these items the whole stage now grows backwards and we need to start earlier of course once the "Plan end" date of the first stage is moved to the right into the future that will now move both the planned start and plan and of the second stage back to the right again for the sake of completeness let's also transform this back into his stage finish to start dependency connection by updating the initial item's did to link to the first stage again make the subsequent items consecutively dependent on that initial one and change all the dependency connection types for the tasks back to finish to start as you see even these fundamental changes in the dependency structure of that schedule can be done within a few seconds that's amazing let's also take a quick look at how to extend the stage with additional items let's add two additional tasks and one milestone to the second stage the first thing you should do in case you have applied the auto stage formulas click into them and adjust the referenced range to also cover these new items same for the workdays to enable the stage to cover the full stage time span for simple consecutive relationships we use the autofill handle then change the milestone dependency connection to finish to finish and enter the number of required workdays for the tasks be aware that this is just a simple example you are totally free to always transform this into a way more complex dependency structure for example we can make this task i4 start together with the initial tasks simply by updating the linked id and switching to a start to start dependency connection now a change in duration of these two tasks not necessarily has an impact on the overall stage duration as these tasks run in parallel now or let's say we want both of these milestones to finish at the same day and figure out what does that implicate for the required start date of the last two tasks let's make this milestone dependent on the other one and these two tasks now dependent on that second milestone with a finish to finish and for this one a start to finish dependency given this setup once we increase the number of workdays here we now just have to start earlier and now this even has an impact on the stage time span all these examples are only a small fraction of all the possible dependency structures you can build so the opportunities with this dependency engine feature are basically endless as quick as we had added these three items within seconds we can also delete them you see the ids on the left perfectly update the reference ranges of the auto stage formulas also do not require a manual update and as we have just deleted a few of these content rows and now only have 16 left at the moment we can easily add new ones anytime just by selecting the last row and using the auto fill feature as soon as we have more potential rows in this gantt chart than fitting on the screen we would need to scroll down in order to see all of them but instead of applying the scrolling to the whole worksheet i recommend to scroll to the top once then select the initial row of the gantt area go to the view tab and click on freeze panes to freeze the header including the separator row now only the gantt area below the header is moved when we scroll and that way it is way more user friendly before we continue let's use this additional space for adding two more stages to this gantt chart we call this one rollout prep and make it dependent on the second stage with a finish to finish dependency connection so this stage is affected by changes in both the second and the first stage and then we add another stage called roll out that will just start right after the rollout prep stage has finished it is simply amazing how quickly the setup of a fully dynamic project plan can be done with this feature as it's incredibly intuitive fast and easy to use and another amazing thing is whenever you're done setting up the dependency structure for your project plan you can just collapse all these columns with one click and that gives you just a clean and dense overview of the most crucial information about the scheduled plan and you still have full control over the workload so adjusting the number of required workdays for an item can be done even in this dense view and anytime you want to make a change in the underlying dependency structure simply reopen the section and you are good to go right at the moment 
-->

</br>
</br>
</br>

## 6. Dynamic Project Timespan

### 6.1. Defining the name references

<p align="justify"> These stages perfectly summarize the duration and time span of their respective tasks and milestones. However, we want to implement a feature that goes one level above that and dynamically keeps track of and visualizes the whole project time span. To do this, we will collapse the extended planning section and focus on the "Plan start" and "Plan end" columns. Using the values in these two columns, we will compute the project time span and duration in workdays. </p>

<p align="justify"> We have already prepared placeholders for the project time span, which is nothing else but the minimum and maximum of these dates. We will not only display these date values but also visualize the time span in the timeline section. To do this, we will create a named calculation called "project_start" that dynamically determines the minimum start date of any item in our project (*). The range that we pass for this starts right at the header row in the "Plan start" column. In order to give this project some space to grow, we go down to row 1000. We will set up a similar named calculation for the "project_end", which is calculated as the maximum of the "Plan end" column (**). </p>

```
(*) project_start = MIN(Gantt!$S$13:$S$1000)
(**) project_end = MAX(Gantt!$T$13:$T$1000)
```

<p align="justify"> For the project duration in workdays, we will create an additional named calculation that is based on the net workdays function which we can call "project_duration". This will calculate the workdays between "project_start" and "project_end" using the NETWORKDAYS function (*). That enables us to replace the manual start date with the dynamically calculated "project_start". We will then replace the placeholder with the "project_end". For the project duration, we will insert a concatenated expression with an opening bracket, then the dynamically calculated project duration, and a closing bracket (**). It is important to mention that this project duration in workdays is based on the project time span and is not necessarily equal to the actual workload, especially if you have tasks and stages that run in parallel. </p>

```
(*) project_duration = NETWORKDAYS(project_start; project_end)
(**) = "(" & project_duration & ")"7
```

![Doc6 1](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/108b9041-7da9-43bd-9a2f-c177de83cf44)

### 6.2. Visualizing the information in the Timeline Area

<p align="justify"> We can now use this information to visualize the project time span in the timeline area. We will make use of the row that is right on top of the current timeline and add a conditional formatting rule to these cells. We want to format all cells for which the date in the timeline is greater than or equal to the "project_start" and smaller than or equal to the "project_end" (*). All the cells for which this condition is true shall get a neat green background fill. That seems to work pretty well, but it still doesn't look perfect because the row still has this standard height, which is a bit over the top. So, we need to make this row much tighter. We will remove this temporary unicode icon description and set a row height to 6 pixels. We will add a similar conditional formatting rule to the actual timeline date rows. The formula for this rule is exactly the same, but for this one, we just make the background a bit darker by changing the fill to this slightly darker gray tone. Both these conditional formattings in combination result in a pretty refined visualization of the project time span. With any change in the schedule, may it be a change in task durations or a change in a dependency structure, this visualization smoothly updates to always capture the full project time span. In addition, this feature becomes even more valuable whenever we have a really large project because once we scroll down and the upper stages are not visible anymore, we still see the overall project time span in relation to the currently visible items, which is pretty amazing. </p>

```
(*) = AND(date>=project_start; date<=project_end)
```

![Doc6 2](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/54fa7011-185f-4ec6-86a7-f399bf017f5c)

### 6.3. Solving a couple of weak points

<p align="justify"> Now that this project time span is dynamically computed, there are still two weak points that we have to take care of. The first weak point becomes apparent when we take a closer look at the initial timeline date cell, because that one is referencing the dynamic "project_start" date visualized in the header. That makes us run into a lot of errors when no dates are provided in a "Plan start" column, because now the "project_start" date value is just 0. To make sure this doesn't happen, we will make the timeline start date more robust than a simple reference to the "project_start" date. We will create a perfectly tailored name calculation called "timeline_plan_start". We will set the scope to this worksheet, and what we want to do in this calculation is first check if "project_start" equals 0 (= IF(project_start=0; ), because then we want to use the TODAY function as the fallback value for the timeline start (;TODAY()). In order to make sure that it always shows the beginning day of the week of today, we simply subtract the weekday number of today with an encoding that returns 0 for Monday up until 6 for Sunday (;TODAY() - WEEKDAY(TODAY();3); ). That way, in case today is a Wednesday, for example, this expression subtracts 2 and returns to Monday of the week. Only in case "project_start" is not zero, then we return this "project_start" value minus its weekday number (; project_start - WEEKDAY(project_start)) with the same encoding. Then we will replace this original formula in this first timeline date cell with the defined named calculation. </p>

```
(*) "timeline_plan_start" = IF(project_start=0; TODAY() - WEEKDAY(TODAY()); project_start - WEEKDAY(project_start; 3))
```

![Doc6 3](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/39d50fa5-ab60-495c-9378-448af9d6e2a4)

<p align="justify"> The second weak point appears whenever we use the auto stage formulas to dynamically calculate the stage time span, but the dates for the items of that stage are not available yet. That will temporarily make the "ind.start" value and thus the "Plan start" and "Plan end" values of that stage 0. It would not be a problem if these values weren't taken into account when calculating the "project_start" date, which is now incorrectly 0 as well. That causes the timeline to use the fallback value as the starting date for no good reason. However, no worries, this issue can be solved with a tiny adjustment in the "Plan start" formula. The initial if statement makes sure that the value in the "ind.start" cell is only taken into account in case the cell is not empty (IF(ind.start<>""; ind.start; ...)). Now we need to change that into making sure that this value is not only non-empty but that it is also greater than zero (IF(ind.start>0...)). This tiny adjustment fully eliminates the second weak point. The only visible impact that we will have is that the autostage workday function now temporarily shows an error, but this is actually a good way of telling us that we need to define a schedule for at least one of the items in that stage before any stage calculation can be done here. Let's update the "Plan start" formula for all the rows. We can perfectly see how the "project_start" date and the timeline show the correct values now, so we are able eventually to redefine the dependencies for the items in the rollout prep stage. </p>

```
(*) Plan start column (final result) = IFERROR(@IF(ind_start>0;ind_start;
                                                  IF(ind_end<>"";plan_start_calculation;
                                                                IF(d.id<>"";
                                                                            IF(OR(d.conn="FS";d.conn="");fs_calc;
                                                                            IF(d.conn="SS";ss_calc;
                                                                            plan_start_calculation));"")));"")
```

</br>
</br>
</br>

## 7. Auto-Coloring Engine with 4 automated color modes

<p align="justify"> After successfully implementing the planning logic, we can now focus on the visualization aspect of the template. We will set up an auto-coloring engine that lets us switch between four different and fully automated color modes. Let's start by hiding the planning section and setting up the drop-down menu for the color mode selection. We will merge 6 cells, set the text alignment to the middle, set the font style to size 10 and bold, and add some gray borders. The label of this drop-down selection will be "Color by" and right below is where we want to put the actual drop-down menu. So, let's also merge the 6 cells below, align it to the middle, set the size to 9 this time, and the background fill to a light color. </p>

<p align="justify"> To add a drop-down selection, we jump straight to the data tab and open the data validation window. We allow a list of values and statically define the list of possible color modes as "Default project structure", "Team roles", and "Issues". In order to make the selected value easily usable in other formulas, we can define a static cell reference in the left corner and call it "color_by". </p>

</br>
</br>

### 7.1. How does the Color generation works now?

<p align="justify"> The gray coloring that we currently have in place for the items in plan in the gantt chart area, is what our default color mode looks like. However, now that we want to implement way more complex and fully automated color modes, we need to introduce a system that lets us represent different coloring logics, but at the same time keeps the formulas in the conditional formatting rules pretty short, easy to understand, and easy to scale. We will build a system of numeric color codes with a "color code generating" formula that will cover the whole logic. Then the numeric color codes returned by this formula for each row will be used to create conditional formatting rules for both the color indicator section and the chart area. We make these color codes referencable by setting up a dynamic name reference called "color_code" for a newly inserted A column called "Color code" (remove de $ sign in front of the row). For now, this color code name simply references the cell content, but later we will insert the final formula directly into the name reference in order to make it a named calculation. </p>

<p align="justify"> As we want all the coloring, including the default colors, to be based on color codes; let's take a quick look at how these default indicator colorings are currently organized in the color indicator column. Well, in the conditional formatting window let's first move the stage coloring rule above the task and milestone rule, and set the stop if true check mark for increased efficiency. For this default coloring, we have two rules in place that simply check what the type of item is. At the same time, we have three different rules for the chart area since the task and milestone items require a different visualization. For our new color code system, we will move these type checks (type of item) over to the color code column formula and let the result be represented by numeric codes. </p>

</br>
</br>

### 7.2. Creating a Color Code System

<p align="justify"> For generating the default color codes in the "Color code" column we start by making sure that the "type" value in row is non-empty, and then we use the SWITCH statement to return different numbers based on which item type is selected. We will keep it simple and return a 1 in case it's a stage, for tasks we're going to return a 2, and for milestones it will be a 3. Otherwise in case "type" is empty we simply return a -1 (*). Let's add this formula to all the rows and that gives us a numeric encoding based on which item type is selected. Let´s copy this formula to create a new name reference called "color_code_default", and substitute the formula in the "Color code" column with the name reference. </p>

```
(*) = IF(type<>""; SWITCH(type; "S"; 1; "T"; 2; "M"; 3); -1)
```

<p align="justify"> Let's open the conditional formatting manager for the indicator color column and change the two coloring rules with this formula. For the stage check rule we replace the logical test with "color_code" equals 1 (*), and for the task and milestone check will be replaced with "color_code" equals 2 or 3 (**). The indicator color column still works as before since the logic behind hasn't changed. So let's do the same for the chart area, we select the chart range, open the conditional formatting manager, and replace the original logical tests with a color code check. </p>

```
(*) = color_code = 1
(**) = OR(color_code=2; color_code=3)
```

<p align="justify"> For the whole formula implemented in the "Color mode" column, we now want to consider the selected color mode and build up the respective formula pattern to act based on it. In order to find out what the currently selected color mode is, we make use of the SWITCH function; enter the "color_by" cell reference as the expression to look at, and, in case the selected mode is "Project structure", we want to execute all the respective calculations for which we enter a placeholder value for now ("ps_cc"). In this SWITCH statement, we can now just use the "color_code_default" calculation as the default return value. Then if "Team rules" is selected as the color mode we want to do a different color code calculation ("tr_cc"), and eventually, in case "Issues" is selected another calculation will be done ("is_cc") (*). Let's update all the other cells with this adjusted formula, when we select a different color mode all the color code cells display the respective placeholder value. </p>

```
(*) = SWITCH(color_by; "Project Structure"; "ps_cc";
                        "Team Roles"; "tr_cc";
                        "Issues"; "is_cc";
                        color_code_default)
```

#### 7.3. "Project structure" color mode

<p align="justify"> For now let's build a color code generating expression only for the "Project structure" color mode. In this mode we're going to make use of 8 main colors from our color palette and we want each stage with all its items to have a new individual main color. We're also going to need different color codes within a stage in order to represent the different types of items with different shades of the same color. </p>

<p align="justify"> As we need to differentiate between stages, we will introduce an automated stage ID system next to the "ID" column. This will automatically assign the same ID to all the items of one stage. To make this stage ID easy to reference, we'll add a dynamic name reference called "Stage ID". So a 1 to all the items of the first stage a 2 for the second stage and so on. To make this stage id easy to reference let's add a dynamic name reference called "stage_id". </p>

<p align="justify"> For the stage ID formula, at first, we need to make sure that "type" is non-empty because a non-defined item cannot belong to any stage. In case a "type" value is available we do the same check we already did for the main ID formula in order to figure out if the item is the first one in the whole list. So in case the row of this cell is two rows below the header cell we know this has to be the first stage and return a 1 for all subsequent items. We want to continue with the same stage id for the tasks and milestone within the first stage, until we reach another stage item for which we then increment the stage id number by 1. That means if "type" equals "S" we compute the maximum id number in the previous column range and then just add 1, since it's a new stage. If "type" is not equal "S" the item is either a task or a milestone, so in that case we just return the maximum id number from the previous column range without adding a 1, since we want all the items in a stage to have the same stage id as its stage. If the row is empty the number displayed would be -1 (*). </p>

```
(*) = IF(type<>""; IF(ROW(stage_id)=ROW($C$14)+2; 1; IF(type="S"; MAX(prev_col_range)+1; MAX(prev_col_range))); -1)
```

<p align="justify"> Now that we have the stage id defined, let's jump right into the color code formula to replace the "Project structure" placeholder with some color code generating expressions. The first thing that we gonna make sure of is, in case the stage id is -1 the color code should also be -1, which means no color will be displayed. However, in case we have an actual stage id available, we are going to build a three-digit color code that we will concatenate as a text first and then transform it into a number. For this "Project structure" mode the first digit for the main color is determined by the stage id, so we use the "stage_id" reference and concatenate it with a text string that will be returned from a SWITCH expression. Based on the selected item type, we either going add "01" for a stage item, for a task item will be "02", and for a milestone item "03". After that, this concatenated text string will then be transformed into a number by wrapping it in the NUMBERVALUE function (*). </p>

```
(*) "ps_cc" = IF(stage_id=-1; -1; NUMBERVALUE(stage_id & SWITCH(Type;"S";"01";"T";"02";"M";"03")))
```

<p align="justify"> We now have a clean numeric encoding for the project structure that beautifully represents both the stage and the item type within each stage. But there is one limitation that we haven't considered yet, we have a limited amount of main colors in our color palette and we want to be able to create an unlimited amount of stages differentiable from the stages immediately around them. So the logical solution for this "Project structure" mode is to reuse the color palette over and over again so that after using all 8 main colors, the ninth color will be the same color as the first one, and so on. </p>

<p align="justify"> For the color code that means, instead of using the stage id directly we're going to use the so-called MOD operation which returns the remainder from dividing the stage id by a divisor. This divisor in our case has to be 8 so this function will basically partitionate the stage id into a part that is divisible by 8 and a remainder that is not divisible by 8 which will then be returned. So in case the stage id is 1 it will return 1 and if it's 9 it will take out 8 because that's divisible by 8 and then return the remainder which is also 1. The only special case that doesn't work in our favor is whenever the stage id is fully divisible by 8 because then it returns the remainder of 0, and in that specific case we actually want to have an 8 as the first digit of the color code. So let's expand the previous formula by an if statement that takes care of the special case. </p>

```
(*) "ps_cc" = IF(stage_id=-1; -1; NUMBERVALUE(IF(MOD(stage_id; 8)=0; 8; MOD(stage_id; 8)) & SWITCH(Type;"S";"01";"T";"02";"M";"03")))
```

<p align="justify"> Let's open the conditional formatting manager for the color indicator column and create two additional rules for the first group of color codes that have a 1 as their first digit. We can simply duplicate the two rules that we already have in place for the default coloring and simply adjust the color codes that these rules are looking for. For the stage item the code is 101, the main color for this first stage is the cyan blue (#4633F2), however, the color selection order is totally up to you. In a similar manner, we adjust the formula for the task and milestone items by changing the rule to 102 and 103, and this time select a lighter shade of the same blue color. Of course we also need to set up the respective rules in the chart area. So let's open the conditional formatting manager, duplicate the three default rules below the stopper rule and change them as we did in the color indicator column. We only need to move the new coloring rules right below that stopper rule. </p>

<p align="justify"> The only reason why the other stages are not visible yet is that we haven't set up the conditional formatting rules corresponding to their color codes. For now let's outsource the "ps_cc" color code creation expression into an own named calculation that we call "color_code_project_structure", and then replace this expression in the main formula using that defined name. </p>

</br>
</br>

### 7.4. "Team Roles" color mode

<p align="justify"> Before we set up the conditional formatting rules for the remaining color code families, first we are going to implement the formula for the "Team roles" color code generation. That's because we want to make the team roles and project structure color modes use the same conditional formatting rules and, thus also, partly, the same color codes. In this context the 8 main colors will represent the 8 team roles that can be defined in the "Settings" worksheet. As we have the team role shortcodes available, we can easily implement a corresponding role id encoding by looking up the role shortcode in the "Settings" worksheet. </p>

<p align="justify"> Let's start by defining a dynamic name reference for this "Role" column, and also define another dynamic name reference for the role id column (remove $ sign in front of the row). Then, in the role id column, we use the MATCH formula to look up the role in the shortcode array (the shortcode column in the "Settings" worksheet). Make sure to make it an absolute range reference with dollar signs and, eventually, enter a 0 to look for exact matches. This statement will simply return the team role's position in the "Settings" team roles column. We can directly use this position as the role id. If the roll cell is empty this will result in an N/A error, so let's add an IFNA statement to catch that error and return a -1 instead (*). </p>

```
(*) = IFNA(MATCH(role; Settings!$I9$:$I16$; 0); -1)
```

<p align="justify"> I still want to make one additional modification to this formula. In my opinion it makes sense to limit the team member assignment to tasks and milestones because these are always connected to real actions, while stages tend to be more abstract concepts mainly used to better structure the whole project and to group items. So my preferred approach, at least for this color-related visualization, is to only return a role id for tasks and milestones while for stages we simply going to return a -1 (*). With this modification you can still assign a team member to a stage if you like but it won't be considered by the auto-coloring engine. </p>

```
(*) = IFNA(IF(OR(type="T", type="M"); MATCH(role; Settings!$I9$:$I16$; 0); -1); -1)
```

<p align="justify"> Let's use this role id to replace the "tr_cc" placeholder in the color codes column with some meaningful computations. The first thing we do here is checking if the role id equals -1; this -1 can have multiple reasons either there is no item defined in the respective row, or there is an item that has no team member assigned or it is a stage item. For all these cases we want to apply the default coloring as fallback option because this ensures that even if there will be no team role-related coloring, the item is still visualized. In case we have an actual role id available then we build a three-digit color code just like we did for the project structure mode. The NUMBERVALUE function again helps us to transform a concatenated text string into a numeric value and within this function we concatenate the role id as the initial digit and then use the SWITCH statement to determine the last two digits based on the item type (*). As there's no need to visualize the hierarchical differences between a stage and its items, we don't need to use different shades of a color this time and thus, we simply use the strong shade version of each main color for both the tasks and the milestones. That means in case the type is a task we generate a color code ending in 01, these are the ending digits that we used for the stage items in the project structure mode, which means we can make use of the exact same conditional formatting rules for the indicator color and gantt chart visualization. For the milestones we have to introduce a new color code ending in 04. With the team roles color mode all color codes will end in either 01 or 04. So let's include the 04 ending into the conditional formatting rules that we already have in place. As we want to have a strong color shade here we're going to make this rule an OR statement that also considers 104, and for the gantt chart area we just include this color code into the rule that we have already set up for the milestones. </p>

```
(*) "tr_cc" = IF(role_id=-1; color_code_default; NUMBERVALUE(role_id & SWITCH(type; "T"; "01"; "M"; "04")))
```

<p align="justify"> What's left for both modes is to set up the conditional formatting rules for all the remaining three-digit color codes that go from 2 to 8 as their starting digit. Before we do that we are going to clean up the color code formula by putting the team roles calculation into a separate named calculation that we'll call "color_code_team_roles". For all the remaining three-digit color codes fortunately the process will be pretty straightforward. In the indicator column we are going to duplicate the color code checking rule seven times, we change the color code (_01 and _04 will have strong colors, while _02 and _03 will display shades of its stage color) and adjust the formatting to a new color. For each new pair of rules we gonna select a new main color from which we picked the strong and light shade versions. The only important thing you have to make sure here is that the color families used for these rules are consistent between the indicator column and the chart area. </p>

</br>
</br>

### 7.5. "Issues" color mode

<p align="justify"> The last color mode will allow the user to highlight items for which an issue has arisen. Since the corresponding formula will be based on the issue column let's make the cells dynamically referenceable as "issue", then copy unicode warnig sign and replace "is_cc" placeholder with a formula. We are going to use an IF statement to instantly check if the "issue" name reference equals the unicode symbol and, if that's the case, we want to generate one of two potential color codes. As for the other two advanced color modes we have only used 1 to eight 8 as the initial digit we now simply create two color codes that are 901 and 902. The number 901 will be used whenever the type is a stage or a task and, as we need a separate rule for the font formatting of milestone symbols, we create a separate code 902 for it. Of course we still want to display all the other items without issues so again, we just use the color code default calculation as the fallback option (*). </p>

```
(*) "is_cc" = =IF(issue="⚠"; SWITCH(type; "S"; 901; "T"; 901; "M"; 902); color_code_default)
```

<p align="justify"> To create the conditional formatting rules for the color indicator column we only need one simple rule that we can use for both color codes so, in case the color code equals 901 or 902 we want to have a warning red fill. For the chart area we need to make the distinction between item types so we duplicate these two rules. The first one will be for the color code 901 and will simply fill the cell with the warning red, while the second rule is meant to be for the milestone symbols (color code 902) and will overwrite the warning sign with a red font style. Finally, we only need to take care of the main color code formula, first we transform the "is_cc" formula into a named calculation called "color_code_issues" and replace that expression. For now own we can hide the column "A" corresponding to the color code generation column we've just built. </p> 

```
Final color code formula = @SWITCH(color_by; "Project Structure"; color_code_project_structure;
                                             "Team Roles"; color_code_team_roles;
                                              "Issues"; color_code_issues;
                                              color_code_default)
```

</br>
</br>
</br>

## 8. Base Plan Comparison

<p align="justify"> Let's move on to the next amazing feature that will help you to take a snapshot of your actual plan, save it as the base plan, and then compare it to the changes that you make to your actual plan over time. The crucial idea of this feature is that once you have built up your schedule with all the dependencies, you can document the date of the recent changes you made. Then, jump over to the base plan section by unhiding its two columns (V and W) and just copy over the actual plan as plain values. Let's decrease the font size and change the font color to a dark gray. Once you have taken that snapshot, you should also document the date up in the section header so that you always know when the snapshot was taken. </p>

<p align="justify"> Let's start by adding another drop-down menu that we call "Display" and that will have two list items to select from. The first option is simply called "Plan" and refers to only showing the actual plan, while the second item will be called "Plan vs. Base" which means we want to have both plans visualized in the chart area at the same time. Let's select "Plan" for now because that is the default option that is already in place here. </p>

<p align="justify"> Then we create a static name for these cells and call it simply "display". And of course, we're also going to add some dynamic name references for both the date columns of the base plan. So for the base start, as always, we're going to remove the dollar sign here. And of course, the base end. Based on this base start and base end, we now want to create a named calculation called "item_in_base" that does exactly the same for the base plan like the "item_in_plan" calculation already does for the actual plan calculation. To remind you, the "item_in_plan" calculation simply checks for every cell in the chart area if the date value in the timeline in the respective column is in between the "Plan start" and "Plan end" date and then returns either true or false. So let's copy the formula, create a new name, call it "item_in_base", set the scope to this worksheet, paste the formula, and now we simply overwrite this with "base_start and "base_end". </p>

```
(*) "item_in_base" = AND(Gantt!date>=Gantt!base_start; Gantt!date<=Gantt!base_end)
```

</br>
</br>

### 8.1. Milestone symbol display adjustment in the Chart area

<p align="justify"> Whenever we select the "Plan vs. Base" option, we now want to use this "item_in_base" calculation to additionally display the original base plan with the actual plan. For the milestone symbols that we currently generate with the formula from within the cells in the chart area, we now have to display an additional milestone symbol earlier. Let's adjust this formula to do exactly that. Instead of simply checking if type equals "M" and ""item_in_plan"", we now want to do this type check initially. After that, we add a second if statement to figure out if the item is in plan in the current cell. If that is the case, we're going to return the milestone symbol independently from the current display mode, because we want the actual plan to always be displayed in both modes. The special aspect to consider now is in case we are in "Plan vs. Base" mode, we only want to display an additional symbol for the base plan if it is not on the same day as the one from the actual plan. So it makes sense to only consider the selected display mode if "item_in_plan" returns false. In that case, we only want to print a milestone symbol for the base plan if the "Plan versus Base" mode is selected and the "item_in_base" calculation returns true. Otherwise, we simply return an empty text string just like we do in case the type does not equal "M" (*). </p>

```
(*) = IF(type="M"; IF(item_in_plan; "◆"; IF(AND(display="Plan vs Base"; item_in_base); "◆"; "")); "")
```

<p align="justify"> Right at the moment, these milestone symbols of the base plan have the default dark blue font color, but we're going to change that into a font color that is less standing out. Let's open the conditional formatting manager and create a new rule that will only target these base plan milestone symbols. The formula for this rule has to test multiple conditions which are: display equals "Plan vs Base", "item_in_base" has to be true, and of course, in case the base and actual plan milestone are on the same date, we don't want this rule to be applied; and we just apply the actual plan coloring. That means we have to put the "item_in_plan" into a not statement to reverse it. And eventually, of course, type has to be "M" (*). For the formatting, I chose the super light gray, which is even lighter than the default coloring of the actual plan. </p>

```
(*) = AND(display="Plan vs Base"; item_in_base; NOT(item_in_plan); type="M")
```

</br>
</br>

### 8.2. Visualizing the base plan Tasks and Stages

<p align="justify"> For visualizing the stages and tasks of the base plan, we will only need to create one conditional formatting rule. Let's select the chart range again, open the conditional formatting manager, and here we can see we should move the milestone rule a bit down, right on top of the actual plan stopper rule. Then we can just duplicate it and adjust this formula for the stages and tasks of the base plan. </p>

<p align="justify"> For these, we also need to make sure that the "item_in_base" calculation returns true. But unlike for the milestones, we want the stages and tasks of the base plan to overlay the actual plan stages and taskbars. If necessary, you will see how that can be done in the formatting options in a second. For the formula, though, this means we don't have to make sure that "item_in_plan" is false. So we can just throw that part out. And eventually, we're going to change the last condition into an OR statement that checks if type is either stage or task (*). </p>

```
(*) = AND(display="Plan vs Base"; item_in_base; item_in_plan; OR(type="S"; type="T")
```

<p align="justify"> For the formatting, we change the font color back to automatic, go to the fill tab, and the reason why we can have both the base plan and the actual plan visualized in the same cell is that we're going to use pattern fills for the base plan. That means if a cell in the chart area is part of both the base plan and the actual plan, the pattern fill of the base plan will simply overlay the full cell fill of the actual plan. As the color for this pattern fill, I recommend using this dark gray because that allows us to use this not-so-heavy point pattern that says "25%" gray. Let's click OK, and there we have the beautiful but still subtle visualization of the base plan compared against the actual plan. </p>

</br>
</br>

### 8.3. Updating the timeline start date in the "Plan vs Base" mode

<p align="justify"> I also want to make sure that the timeline will automatically adjust its initial date to perfectly show the full extent of both the actual and the base plan. Right now, the initial timeline start date is only based on the actual plan. For that reason, the base plan, which starts on the 7th of May, is not fully visible here. In order to change that, let's start by adding a named calculation that computes the project start date according to the base plan, which is simply the minimum value within the base start column. Using the start date of the base plan, we can then set up another calculation similar to the "timeline_plan_start" calculation. We're going to call this "timeline_base_start," and it will take the raw project base start and calculate the start of the respective week by subtracting its WEEKDAY number, which will be encoded as from 0 for Monday to 6 for Sunday. So this will return the Monday of the week of the base plan start (*). </p>

```
(*) timeline_base_start = project_base_start - WEEKDAY(project_base_start; 3)
```

<p align="justify"> Now that we have both the "timeline_plan_start" and the "timeline_base start defined, we can now make the initial timeline date dependent on the display mode. In case display equals "Plan vs Base," then we have to make sure that the timeline base start is greater than zero. If that is true, we're going to return the earlier of both the timeline base start and the "timeline_plan_start". That makes sure that no matter if the base plan is ahead or behind the actual plan, the timeline will always optimally adjust (*). In case we haven't taken a snapshot yet, which means the "timeline_base_start" is not greater than zero, the "timeline_plan_start" will be directly used even if the "Plan vs Base" mode is active. Furthermore, in case the selected display mode is not "Plan vs Base," then the "timeline_plan_start" will be used just like before. Once we hit enter, you can instantly see how the timeline perfectly adjusts to show the full base plan. </p>

```
(*) = IF(display="Plan vs Base"; IF(timeline_base_start>0; MIN(timeline_base_start; timeline_plan_start); timeline_plan_start); timeline_plan_start)
``` 

</br>
</br>
</br>

## 9. Progress Tracking

### 9.1. Progress tracking Engine

<p align="justify"> Next, let's add an intuitive and partly automated way of tracking progress and visualizing it. For this, we have one column left at the end of the input section. We begin by changing the format of the numbers to percentages. That way, we can enter the percentage of completion for each item, and it automatically gets the percentage symbol assigned. Every percentage of completion that we have entered should then be reflected in the Gantt chart accordingly. To activate and deactivate the Progress visualization, let's add another drop-down menu with the label "Show Progress" and "Yes" or "No" as the selectable options. To easily access the selected value in other formulas, let's call this cell "show_progress". Of course, we also need this percentage completed value accessible as well, so we create a dynamic name reference called "percentage_complete."

<p align="justify"> Taking this "percentage_complete" value, we then want to visualize the corresponding number of completed workdays in the chart area. To figure out whether each cell in the chart area is part of this completed time span or not, we need to set up a name calculation that is similar to the "item_in_plan" calculation. So let's copy that formula, create a new name called "item_in_complete", set the scope as usual, and paste the formula. Because we want the progress bar to start from the beginning of an item, we can leave the first condition that makes sure that the date in that column is greater equal to "plan_start". What we have to modify is the second condition because instead of just taking the "plan_end" as the upper limit, we want to dynamically compute the end of the completed time span starting from the planned start and then adding the equivalent number of completed workdays. The function that matches best our needs is the WORKDAY function, because it allows us to enter the "plan_start" and then add the number of completed workdays. To compute the number of completed workdays, we simply multiply the percentage complete and the number of required workdays, and since the "plan_start" is already one of the days in the time span, we have to subtract one from this product (*).

```
(*) item_in_complete = AND(Gantt!date>=Gantt!plan_start; Gantt!date<=WORKDAY(plan_start; percentage_complete*wworkdays-1))
```

</br>
</br>

### 9.2. Progress tracking visualization of Stages and Tasks

<p align="justify"> Now we can use it to set up conditional formatting rules for the progress visualization. The first rule will be targeting the stage and task items. We want to highlight a cell as completed workday whenever we have selected "Yes" in the "Show Progress" drop-down selection, and the column date is within the time span of completed workdays. So "item_in_complete" has to be true, and since this rule is for stage and task items, the type has to either equal "S" or "T". If all these conditions are met, we want the cell to be filled with the same dark blue color of the header (*). In the conditional formatting manager, it is now crucial that this rule is placed somewhere on top of the other actual plan coloring rules because that makes sure that this dark blue progress fill will always cover and not be covered by all the subsequent rules. For increased efficiency, don't forget to set a checkmark at the "Stop If True" option.

```
(*) = AND(show_progress="Yes"; item_in_complete; OR(type="S";type="T"))
```

</br>
</br>

### 9.3. Progress tracking visualization of Milestones

<p align="justify"> Now that we have taken care of the stage and task items, let's focus on the progress visualization for the milestones. First, we're going to create a conditional formatting rule for the coloring, and for that, we can simply duplicate the progress rule that we have just set up, move it down one time, and then adjust this last condition to "Type" equals "M". For the formatting, we don't want to have a cell fill, but we want to set the font color to the same dark blue. That's it. So now, as soon as that percentage value is set to 100, it's now colored as completed, which is great. But to make this even greater, let's make the milestone symbol also transform into this finish flag icon to emphasize that we have successfully crossed the finish line. </p>

So let's copy it and modify the symbol generating formula. As we want to have the progress only visualize for the actual plan and not the base plan, this is the only place we're going to do any modifications. Instead of directly printing this milestone symbol, we now gonna ask if "Show Progress" is set to "Yes" and if maybe this milestone is already completed. So if "Item Incomplete" is true, we print that beautiful flag symbol. Otherwise, just a standard milestone symbol. Let's close that if statement and add this updated formula to the whole Gantt chart range.

And there is the beautiful finish flag for this completed milestone. Amazing! When we delete that 100, the symbol perfectly transforms back to the regular milestone symbol, and putting it back in there will just make it re-transform. Of course, this milestone transformation works for every single milestone, and the whole progress visualization can be turned on and off up there.

So it is safe to say that the fundamental Proquest tracking feature is completed by now, and that gives us the great opportunity to take a closer look at how we can partly automate the progress tracking itself. What makes most sense to automate here is the Proquest tracking of the stages because for most use cases, it makes sense to have the progress of the stage item linearly reflect the progress over all items or at least the tasks covered in that stage.

For that purpose, I have prepared another auto-stage placeholder formula up there. This formula basically just computes the sum product of the workdays and the progress of all items belonging to this stage and then divides that by the sum of total work days over all items at this stage. Once applied, the stage progress now automatically updates every time we change the progress value for these tasks. Once all the items in the stage are completed, the stage is completed as well, and that just saves you a lot of manual work.

To highlight the fact that the stage progress is linked to other cells, we're going to use the same blue font color formatting that we already have set up for multiple other columns like the workdays column, for example. Here we have that respective conditional formatting rule. Let's add this column to the relevant range by typing in a comma and selecting the column values. Perfect! And now we instantly see, okay, this stage progress is automatically calculated and linking to other cells.

Let's add this auto-stage progress formula to the other stages as well. And keep in mind, if you have a different number of items in a stage, you have to adjust the reference ranges in that formula to make it work correctly. In our example case here, where all the stages have the same number of items, we were able to automate the progress tracking of all these stages within a few seconds.

Regarding the milestone items here, there is one thing that you should be aware of. With this auto-stage progress formula, you're going to create references to the workday and Proquest cells directly and not to their dynamic name references, which also means an empty workday cell like we have here for the milestone item will be counted as zero and not as one and thus not impact the calculated stage progress. At least not until we put an actual one in there.

In my personal opinion, that actually is a good thing because milestones are something that should not be measured in workdays, and they are often tied to the completion of tasks anyway. So I will leave the workday cells empty. In case you want your milestones to contribute linearly to the stage progress, all you need to do is just enter a one in that workday cell.

Since I just mentioned that the completion of milestones might be tied to the completion of tasks, let me quickly show you how that can be done correctly. Let's say that we want this milestone to be set to 100 percent only in case that this task here is fully completed. For that, we can type in the formula "1 times that progress value equals 1." This whole expression will only return one, or in other words, 100 percent, in case the task completion is exactly 100 percent. In all other cases, this will always return zero. Once we hit enter, we see that the milestone progress is now highlighted in blue, telling us that this links to another cell. And with the task being at 100 percent completion, the milestone is completed as well. But as soon as we only take away one percent here, the milestone completion is set to zero percent. Of course, this concept can also be extended to multiple tasks. For example, to make the milestone completion dependent on the completion of all task items in the stage, we can check "if the sum of these values equals the number of task items." And now, whenever any of these tasks is not completed anymore, the milestone completion is immediately set back to zero.

I think this concept alone is something that makes this Proquest tracking and visualization feature even more powerful. At this point, we have achieved a huge milestone as all the formulas and conditional formatting rules for visualizing stages, tasks, and milestones in any possible form are now completed. The whole implementation is already pretty lightweight, efficient, and reactive. Yet there's still some improvement potential, which you might recognize when we turn on and off the progress visualization.

The progress bars and flags appear and disappear pretty fast, but there is still a tiny, tiny lag of a few milliseconds, which we can decrease a lot with a powerful trick. And that trick I'm going to reveal to you right now. One big reason why at the moment it takes a bit longer for this worksheet to update when we, for example, turn on the Proquest visualization is that the visualization of the milestone symbols is a two-step process. The first step in this process is to print the actual milestone symbols into the cell based on this in-cell formula. And then in a second step, we apply the conditional formatting rules to give them their respective color.

It is not possible to entirely eliminate this first step as the cells still need to have some content to display these Unicode icons. Yet we are able to simplify the cell content by making the icon definition part of the respective conditional formatting rules. To do that, let's copy this default milestone icon, open the conditional formatting manager, and scroll straight to the bottom to start with the first milestone coloring rule. At the moment, this rule only defines the color for this milestone icon. But when we go to the number tab and select the custom category, we can apply the same trick we previously used for the auto-stage start date.

Instead of defining a number pattern, we just paste in the milestone icon here. At the moment, the sample section shows nothing because to make this work, we need to have some numeric value written into the cell, which we will do in a second. But for now, let's click OK, and we instantly see that this milestone symbol is now the default number formatting applied by this rule.

Let's quickly adjust the formatting for all the other milestone rules in the same way. This also includes the issues milestone rule and the one to display the base plan milestones. Only for the Proquest milestone rule, we're going to need to use the flag icon. So let's copy this one, reopen the conditional formatting rule manager, and adjust the number format for this rule accordingly.

</br>
</br>
</br>

## 10. Date Highlighting (Dynamic & Static)

<p align="justify"> To perfectly complement the ProQuest tracking feature, we are now introducing a new feature that allows you to highlight a specific day in the timeline. You can either statically highlight a particular day or dynamically highlight the current day to put the project progress into perspective. For this date-highlighting feature, we have created a new drop-down selection that has all the dates in the timeline and an additional option for today. The drop-down selection has the label "Highlight," and the static reference name is "highlight." We changed the background fill to a light green. This list dynamically updates and always displays exactly those dates that are visible. Although the date cells spread over multiple rows, we get a warning alert because we can only reference either a single row or a single column. We manually adjusted this to only reference the cells in row 10, making all the values accessible. </p>

<p align="justify"> To highlight the selected date in the chart area, we created a new conditional formatting rule that is able to work with both the dynamic and the static highlighting options. We used an if statement to check if the "highlight" equals "today". In that case, we compare the date in the timeline with the TODAY function, which returns an actual date. Alternatively, "highlight" has to be one of the static dates in which case we can directly compare the date and the "highlight' value (*). When it comes to formatting, we chose a dark blue right border with the dotted style, which gives us a continuous line positioned right at the end of the selected date. For the timeline date and weekday section up there, we added a similar conditional formatting rule that has the same formula. We chose a dark blue fill and a bold and green font style. The selected date is now beautifully highlighted in the timeline section while the respective dotted line is positioned right at the end of that highlighted day. This setup allows us to switch between any of the visible timeline dates or use the dynamic today option. </p>

```
(*) = IF(highlight = "Today";  date = WORKDAY(TODAY();-1;1);  date = highlight)
```

</br>
</br>
</br>

## 11. Scroll Buttons

<p align="justify"> As the final step, I want to show you how to add some easy-to-use scroll buttons to dynamically move the timeline to the right and left. This will be especially helpful if you have a larger project that doesn't fully fit the screen at once. The way scrolling for this timeline works is actually pretty simple. To scroll to the right, we only need to add a number of workdays to the initial timeline start date, and as we want to control this number of added workdays using scroll buttons, we need to put that value somewhere in the worksheet. Let's just use the cell on the left side of the timeline star date, put a zero in here for now, and give the cell a name reference called "scroll_increment". Now we can add that value to the timeline start date via the WORKDAY function (*). </p> 

```
(*) = WORKDAY(timeline_start; scroll_increment)
```

<p align="justify"> To intuitively control this number, let's go to the "Developer" tab, which you can activate in the Excel settings in case it is not visible, and then we insert the form control element called "scroll bar." Put it in the right-superior corner of the timeline area and adjust its size so that only the scroll buttons are available. To make these buttons control the "scroll_increment" number, we simply have to link it to the "scroll_icrement" cell. Then we set the maximum value to 50, let the incremental change at 1 for now, and remove the 3D shading for a cleaner design. To increase the scroll speed, we could either modify the formula in the initial timeline date cell or alternatively just right-click and format these buttons to increase the value for the incremental change to 5, for example. That's all we need to make this timeline scroll to the right and left using our mouse only. </p>

![Doc11](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/d4be8a7c-947a-461c-964e-13685ebebc5a)

</br>
</br>
</br>

## 12. Manage & Filter Items (Via Quick Access Toolbar)

<p align="justify"> To make the most out of this template, I'd like to give you some additional advice on how to add, manage, and filter items in this Gantt chart. The most common way of adding new items is simply adding them at the bottom. If you want to add multiple items, make sure that you have one empty row available that is part of the Gantt chart. Then you can easily select the whole row and use the autofill function to add as many new rows to the Gantt chart as you want with the same conditional formatting rules, drop-down menus and formulas. If you want to insert an item between other items, for example, I recommend using the quick access toolbar. This is the section in your Excel application that, by default, includes save, undo, and redo commands. You can add more commands, these are the ones I consider useful for a better user experience: fill down, fill up, insert rows, and filter. You can customize the quick access tool-bar ribbon in the Excel Settings, simply select the command on the left side, click on add, and potentially change the order of these commands. </p>

![Doc12 1](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/7957fa5b-28e1-4d83-95a7-17cf60312848)

</br>
</br>

### 12.1. Filtering

<p align="justify"> In addition, the way we have set up this Gantt chart offers a lot of exciting dimensions to filter the items by. For quick and lightweight filtering, I recommend making use of the native filter function that you can easily include in your quick access toolbar to filter the Gantt chart by a specific column. Just select the respective column, including the column header, then press the filter command, which adds this filter arrow to the column header. This allows you to do your filter selection. </p>

<p align="justify"> For example, we can easily create a high-level view of the project by displaying stages only, which gives us a compressed view of the full schedule and the general project progress. Removing the filter is as easy as pressing the filter command one more time. Since we have defined a stage ID in an own column, we can use the filter command to only display selected stages, including their task and milestone items. That allows us to have multiple stages that are actually placed far apart from each other in a condensed view. Another great idea is to filter by the assigned team role, which allows you to break down the project not simply based on single team members, but instead based on areas of responsibility. Two other really interesting columns are the date columns because when applying a filter to one of these, you can choose from a huge selection of specific date filters. </p>

![Doc12 2](https://github.com/AlvaroM99/Excel---Gantt-Chart-for-Project-Management/assets/129555669/5df66ed7-04bc-4d49-8f51-1a14e76d104b)

</br>
