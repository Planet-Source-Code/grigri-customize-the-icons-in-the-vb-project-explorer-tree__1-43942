VBIcons: IDE
JP Mortiboys, 2003
v 1.0 (if that)
Written on VB6/XP, should work on most/all windows platforms, perhaps with minor adjustments.
----------------------------------------------------------------
Purpose:

Allows customization of the icons displayed in the VB Project Explorer window.
Icons can be customized on a global level (eg all form icons) or on an individual per-component level.
I personally use it for visually separating collection classes, interfaces and normal classes.
It allows 32-bit icons to be used, as opposed to VB's standard 16-colour ones.

Known issues:

* Because the addin uses windows messages involving pointers, it must execute from within the VB process. This means it will only work in compiled mode, you can't run it from the IDE.
* Changing the view in the VB project explorer from foldered to nonfoldered causes a complete refresh of the tree, and customized icons are lost until the project is reloaded.
* The UI looks shite.

Future enhancements planned:

* Icon handling for Related Documents
* Customizable category folders
* The option to display only the Component name in the treeview (no filename)
* User-defined sorting of components, drag-n-drop, ...