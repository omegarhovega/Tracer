# :beginner: Tracer
Excel Add-in for fast and clear trace of formula dependents and precedents :rocket:


# :notebook_with_decorative_cover: Installation
- Download and place TraceFormulas.xlam file in C:\Users\Admin\AppData\Roaming\Microsoft\AddIns or equivalent folder
- In Excel, select Options > Add-ins and navigate to Excel Add-ins using the drop-down menu and Go... button at the bottom
- Put a check next to Traceformulas in the appearing Add-ins window, then click OK to close
- To quickly access the Tracer Add-in functions, right click on the Ribbon and choose "Customize the Ribbon"
- Under "Choose commands from" select "Macros" in the drop-down
- You will then see the FindAndDisplayDependents and FindAndDisplayPrecedents functions which can be added to a custom group within a Ribbon Tab of your choice
- For even faster access, add these functions then to the Quick Access Toolbar 

# :sunglasses: Usage
- When calling either FindAndDisplayDependents or FindAndDisplayPrecedents, a window opens showing the reference cell and all precedent or dependent cells with their addresses and their values
- Navigate through the precedents or dependents using up and down arrows on the keyboard or the mouse, the cell will automatically jump to the selected entry
- To exit the trace window and stay on the currently highlighted cell, press 'Escape' or klick the 'Close' button
