# createPowerPointSlides
Automatically create a PowerPoint file using the contents of a TXT file.

Script written by Eleni Christoforidou in MATLAB R2022b.

This script creates a PowerPoint file using the contents of a TXT file called "slideContent.txt" located in the working directory. The file is saved as "newPresentation.pptx" in the working directory.

The slide titles begin with the phrase "Slide X: " in the TXT file, and everything else on the following line(s) is considered to be the slide content, until another "Slide X: " is encountered. The "Slide X: " phrase is not included in the actual slide title. 
