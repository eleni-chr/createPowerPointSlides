%% Script written by Eleni Christoforidou in MATLAB R2022b.

% This script creates a PowerPoint file using the contents of a TXT file
% called "slideContent.txt" located in the working directory. The file is
% saved as "newPresentation.pptx" in the working directory.

% The slide titles begin with the phrase "Slide X: " in the
% TXT file, and everything else on the following line(s) is considered to
% be the slide content, until another "Slide X: " is encountered. The
% "Slide X: " phrase is not included in the actual slide title. 

%%
% Load the Report Generator API
import mlreportgen.ppt.*

% Read the slide content from the TXT file
filename = 'slideContent.txt';
file = fopen(filename, 'r');
fileContent = fscanf(file, '%c', inf);
fclose(file);

% Split the content into slides using the keyword "Slide X: "
slideData = regexp(fileContent, 'Slide\s+\d+:\s+', 'split');
slideData(1) = []; % Remove the first element since it is empty

% Create a new PowerPoint presentation
ppt = Presentation('newPresentation.pptx');
open(ppt);

% Loop through the slide data to create slides
for i = 1:length(slideData)
    slideText = slideData{i};
    slideLines = strsplit(slideText, '\n');
    slideLines = slideLines(~cellfun('isempty', slideLines)); % Remove empty lines
    
    % Create a new slide
    slide = add(ppt, 'Title and Content');
    
    % Set slide title
    replace(slide, 'Title', slideLines{1});
    
    % Set slide content
    replace(slide, 'Content', strjoin(slideLines(2:end), '\n'));
end

% Close and save the PowerPoint presentation
close(ppt);
