 ===================================================================================
 ===================================================================================
 Selections test code ver 1.0

 by M Ferris - Intact Interactive Software.
 url   : http://www.intactinteractive.com
 email : mferris@zfree.co.nz

 ===================================================================================
 ===================================================================================

 This sample code has been given to the development community by Intact Interactive
 software. The code is not finished commercial level stuff so you will find bugs and
 quirks in it I'm sure. Neither is the code a complete application, it is merely a
 testbed for the functionalily being tested!
 So you use this code on an as is basis with no warranty given or implied by Intact
 Interacive. If you are cool with this then go for it, if not, then delete this.

 Also remember that this is not a one way street, I have to survive as a developer
 as well, so obviously I am not giving away all my knowledge. And I would appreciate
 it if you gave me feedback, vote for me at PSC or wherever you found this posted
 and, of course, download the eval versions of my commercial stuff, if you like what
 you see you might even consider buying some of it ...

 A more complete example of this sample source code can be found at
 http://www.intactinteractive.com so I encourage you to visit our site and check it
 out as well as our other source code samples ...

 Some notes about this version.
 ------------------------------

 Features :

            - multiple undo/redo levels.
           - progressive mask building, i.e. you can add/subtract in multiple steps
           - add/subtract mode of mask building
           - rectangle,ellipse and polygon selection tools functional

 How to use :
              - Click on a selection tool to use it.
              - For rectangle and ellipse tool - click on canvas and drag to
              define your selection.
              - For polygon tool - click to start a line, move to endpoint and
              click again to end the line and begin next line, to close the shape
              click near to the first lines starting point.
              - Click the plus button for additive mask, and the minus for
              subtractive mask mode.
              - When undo is possible, the undo button will appear, click it to
              undo to the last level.
              - When redo is possible, the redo button will appear, click it to
              redo to the last level.
              - To fill the selection with red paint - click the paint bucket button.

 To Do :
           - implement the paintbrush selection tool
           - implement the wand selection tool
           - implement disabled button states for undo/redo
           - implement button states for pressed and over on all buttons
           - modify code to work as an activex class (dll) and activex control (ocx)
           - combine code with gradients sample to show implementation of area
           - filling ...
           - add combine mode for xor - i.e. removes the unions
           - add combine mode for and - i.e. keeps only the union
           - add mask transparency functionality
           - add mask feather functionality

 ===================================================================================
 this form illustrates the use of regions to create selection areas for a paint
 program, it shows how to use CombineRgn() to progressively add to a mask region,
 how to use GetRegionData() to implement an undo/redo feature, and generally shows
 how the use of regions can make the implementation of complex selections possible
 in VB !

 the selection tools included in this sample are :

                                   rectangular selection
                                   elliptical selection
                                   magic wand selection (yet to be implemented)
                                   ploygonal selection
                                   paint brush selection (yet to be implemented)

 also an undo function is available to allow the mask to be undone and redone up
 to all levels of the mask creation ...
 ===================================================================================
