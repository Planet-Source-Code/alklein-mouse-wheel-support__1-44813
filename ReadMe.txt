Ben Jones did the real work, figuring out how to sub-class a control and how to capture the mouse wheel message.

I have a program with a few child forms, with grids on some of them, so I needed to scroll the grids with the mouse wheel, but only the grid that's showing on the top form, so I had to add multiple forms handling.

Since I still can't decide whether to allow scrolling only when the focus is on the grid, or if it's anywhere on the form, (I'll leave it up to the customer) I included both ways in this example.  Which one you get is determined by whether you sub-class the control or the form itself.

I hope someone finds this information useful.

Al Klein