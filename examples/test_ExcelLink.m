% Test for ExcelLink
% Author: Milos D. Petrasinovic <mpetrasinovic@mas.bg.ac.rs>
% Structural Analysis of Flying Vehicles
% Faculty of Mechanical Engineering, University of Belgrade
% Department of Aerospace Engineering, Flying structures
% https://vazmfb.com
% Belgrade, 2022
% ---------------
%
% Copyright (C) 2022 Milos Petrasinovic <info@vazmfb.com>
%  
% This program is free software: you can redistribute it and/or modify
% it under the terms of the GNU General Public License as 
% published by the Free Software Foundation, either version 3 of the 
% License, or (at your option) any later version.
%   
% This program is distributed in the hope that it will be useful,
% but WITHOUT ANY WARRANTY; without even the implied warranty of
% MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
% GNU General Public License for more details.
%   
% You should have received a copy of the GNU General Public License
% along with this program.  If not, see <https://www.gnu.org/licenses/>.
%
% ---------------
close all, clear all, clc, tic
disp([' --- ' mfilename ' --- ']);

addpath([pwd '\..\']);

xlApp = ExcelLink(1); % Run Microsoft Excel prgoram
% xlApp.Open([pwd '\test.xlsx']); % Open existing file
xlApp.New; % Create new file

% Add new sheet
xlApp.AddSheet('Test');

% Remove sheet
xlApp.RemoveSheet(2);

% Rename sheet
xlApp.setSheetName(1, 'Test2');

% Write data
xlApp.write(1, [1, 2, 1, 2], [2, 2, 3, 3], ...
  {'test', 'test2', '=5', '=2.454'});

% Read data
data = xlApp.read(1, 2*ones(1, 5), 2:6);
disp(data);

% Change column widht and row height
xlApp.setColumnWidth(1, 1, 50);
xlApp.setRowHeight(1, 2, 30);

% Cell formating
xlApp.format(1, 1, 2, 'CharBold', true, 'CharFontSize', 16, ...
  'HorizontalAlignment', 'Center')
xlApp.format(1, 2, 2, 'Borders', {'Thick', [255, 0, 0], []}, ...
  'HorizontalAlignment', 'Right', 'CharFontName', 'Arial', ...
  'CharColor', [0, 255, 0])
xlApp.format(1, [2, 3], [4, 4], 'CellInteriorColor', [255, 0, 0], ...
  'TopBorder', {'Medium', [0, 0, 255], 'Dash'})

xlApp.SaveAs([pwd '\test.xlsx']); % Save document
% xlApp.Quit(); % Close program

% - End of program
disp(' The program was successfully executed... ');
disp([' Execution time: ' num2str(toc, '%.2f') ' seconds']);
disp(' -------------------- ');