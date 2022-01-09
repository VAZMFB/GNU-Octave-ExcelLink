% Class ExcelLink
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
classdef ExcelLink < handle
  properties (Access = public)
    ServiceManager;
    Document;
    Sheets;
    XlBordersIndex;
    XlBorderWeight;
    XlLineStyle;
    XlHAlign;
    XlVAlign;
    XlOrientation;
  end
  properties (Access = private)
    flag;
    args;
    opened = 0;
    SheetNum = 0;
  end
  methods
    function self = ExcelLink(mode)
      % mode - flag for program visibility
      
      if(nargin ~= 1 || (mode ~= 1 && mode ~= 0))
        mode = 1;
      end
      
      self.flag = (exist('OCTAVE_VERSION', 'builtin') > 0);
      if(self.flag)
        pkgFname = 'windows';
        min_version = '1.6.0';
        fpkg = pkg('list', pkgFname);
        if(~isempty(fpkg)) 
          if(nargin > 1)
            if(compare_versions(fpkg{1}.version, min_version, '>='))
              if(~fpkg{1}.loaded)
                pkg('load', pkgFname);
              end
            else
              disp(['Wait for ' pkgFname ' package to be updated...']);
              pkg('update', pkgFname);
              pkg('load', pkgFname);
              disp(' Package is updated and loaded...');
            end
          else
            if(~fpkg{1}.loaded)
              pkg('load', pkgFname);
            end
          end
        else
          disp(['Wait for ' pkgFname ' package to be installed...']);
          try
            pkg('install', '-forge', pkgFname);
            pkg('load', pkgFname);
            disp(' Package is installed and loaded...');
          catch
            error('Package installation failed!');
          end
        end

        warning('off', 'Octave:data-file-in-path');
      end
      try
        self.ServiceManager = actxserver('Excel.Application');
        set(self.ServiceManager, 'Visible', mode); % visibility
      catch
        error('Could not start Microsoft Excel!');
      end
      
      % Enumerations
      self.XlBordersIndex.DiagonalDown = 5;
      self.XlBordersIndex.DiagonalUp = 6;
      self.XlBordersIndex.EdgeBottom = 9;
      self.XlBordersIndex.EdgeLeft = 7;
      self.XlBordersIndex.EdgeRight = 10;
      self.XlBordersIndex.EdgeTop = 8;
      self.XlBordersIndex.InsideHorizontal = 12;
      self.XlBordersIndex.InsideHorizontal = 11;
      
      self.XlBorderWeight.Hairline = 1;
      self.XlBorderWeight.Medium = -4138;
      self.XlBorderWeight.Thick = 4;
      self.XlBorderWeight.Thin = 2;
      
      self.XlLineStyle.Continuous = 1;
      self.XlLineStyle.Dash = -4115;
      self.XlLineStyle.DashDot = 4;
      self.XlLineStyle.DashDotDot = 5;
      self.XlLineStyle.Dot = -4142;
      self.XlLineStyle.Double = -4119;
      self.XlLineStyle.LineStyleNone = -4118;
      self.XlLineStyle.SlantDashDot = 13;
      
      self.XlHAlign.Center = -4108;
      self.XlHAlign.CenterAcrossSelection = 7;
      self.XlHAlign.Distributed = -4117;
      self.XlHAlign.Fill = 5;
      self.XlHAlign.General = 1;
      self.XlHAlign.Justify = -4130;
      self.XlHAlign.Left = -4131;
      self.XlHAlign.Right = -4152;
      
      self.XlVAlign.Bottom = -4107;
      self.XlVAlign.Center = -4108;
      self.XlVAlign.Distributed = -4117;
      self.XlVAlign.Justify = -4130;
      self.XlVAlign.Top = -4160;
      
      self.XlOrientation.Downward = -4170;
      self.XlOrientation.Horizontal = -4128;
      self.XlOrientation.Upward = -4171;
      self.XlOrientation.Vertical = -4166;
    end
    
    function Visible(self, mode)
      % mode - flag for program visibility
      
      if(self.opened)
        if(mode == 1 || mode == 0)
          set(self.ServiceManager, 'Visible', mode);
        end
      else
        error('No opened document!');
      end
    end
    
    function Close(self)
      if(self.opened)
        
        invoke(self.Document, 'close', true);
        self.opened = 0;
      else
        error('No opened document!');
      end
    end
    
    function Quit(self)
      if(self.opened)
        invoke(self.Document, 'close', true);
      end
      invoke(self.ServiceManager, 'quit');
      delete(self.Sheets);
      delete(self.Document);
      delete(self.ServiceManager);
      delete(self);
    end
    
    function Open(self, filePath)
      % filePath - full file path
      
      if(~self.opened)
        try
          self.opened = 1;
          self.Document = self.ServiceManager.Workbooks.Open(filePath);
          self.Sheets = self.ServiceManager.Worksheets;
          self.SheetNum = self.Sheets.Count;
       catch
        self.opened = 0;
        error('Could not open document!');
       end
      else
        disp('Document already opened!');
      end
    end

    function New(self)
      if(~self.opened)
        try
          self.opened = 1;
          self.Document = self.ServiceManager.Workbooks.Add;
          self.Sheets = self.ServiceManager.Worksheets;
          self.SheetNum = self.Sheets.Count;
       catch
        self.opened = 0;
        error('Could not create new document!');
       end
      else
        disp('Document already opened!');
      end
    end
    
    function AddSheet(self, name)
      % name - new sheet name
      
      if(self.opened)
        if(ischar(name))
          flag = 0;
          try
            self.Document.Sheets(name);
            flag = 1;
          catch
            self.SheetNum = self.SheetNum+1;
            self.Sheets.Add.Name = name;
          end
          if(flag)
            error('Sheet with this name already exists!');
          end
        else
          error('Could not add sheet!');
        end
      else
        error('No opened document!');
      end
    end
    
    function RemoveSheet(self, sheet)
      % sheet - sheet index number
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(self.SheetNum == 1) 
            error('Workbook must conatin at least one sheet!');
          else
            self.Sheets.Item(sheet).Delete;
            self.SheetNum = self.SheetNum-1;
          end
        else
          error('Sheet does not exists!');
        end
      else
        error('No opened document!');
      end
    end
    
    function setSheetName(self, sheet, name)
      % sheet - sheet index number
      % name - new sheet name
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(ischar(name))
            flag = 0;
            try
              self.Document.Sheets(name);
              flag = 1;
            catch
              self.Sheets.Item(sheet).Name = name;
            end
            if(flag)
              error('Sheet with this name already exists!');
            end
          else
            error('Could not add sheet!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function write(self, sheet, col, row, data)
      % sheet - sheet index number
      % col - cell column index number
      % row - cell row index number
      % data - data for write operation
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(length(col) == length(row))
            if(~iscell(data))
                data = {data};
            end
            Sheet = self.Sheets.Item(sheet);
            coli = self.getColumnLetter(col);
            for i = 1:length(row)
              if(isnumeric(col) && isnumeric(row) && ...
                  all(col > 0) && all(row > 0))
                Cell = Sheet.Range([coli{i}, num2str(row(i))]);
                Cell.Value = data{i};
              else
                error('Invalid cell position!');
              end
            end
          else
            error('Different dimensions of column and row vectors!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function data = read(self, sheet, col, row)
      % sheet - sheet index number
      % col - cell column
      % row - cell row
      % data - data after read operation
      
      data = {};
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(length(col) == length(row))
            Sheet = self.Sheets.Item(sheet);
            coli = self.getColumnLetter(col);
            for i = 1:length(row)
              if(isnumeric(col) && isnumeric(row) && ...
                  all(col > 0) && all(row > 0))
                Cell = Sheet.Range([coli{i}, num2str(row(i))]);
                data{i}.Value = Cell.Value;
                data{i}.Formula = Cell.Formula;
              else
                error('Invalid cell position!');
              end
            end
          else
            error('Different dimensions of column and row vectors!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function setColumnWidth(self, sheet, col, width)
      % sheet - sheet index number
      % col - column index number
      % width - coulumn width
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(isnumeric(col) && isnumeric(width) && ...
               all(col > 0) && all(width > 0) && ...
               length(col) == length(width))
            Sheet = self.Document.Sheets(sheet);
            coli = self.getColumnLetter(col);
            for i = 1:length(col)
              Sheet.Columns(coli{i}).ColumnWidth = width(i);
            end
          else
            error('Could not set row height!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function setRowHeight(self, sheet, row, height)
      % sheet - sheet index number
      % row - row index number
      % height - row height
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(isnumeric(row) && isnumeric(height) && ...
               all(row > 0) && all(height > 0) && ...
               length(row) == length(height))
            Sheet = self.Document.Sheets(sheet);
            for i = 1:length(row)
              Sheet.Rows(row(i)).RowHeight = height(i);
            end
          else
            error('Could not set row height!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function format(self, sheet, col, row, varargin)
      % sheet - sheet index number
      % col - cell column index number
      % row - cell row index number
      % varargin - additional variables
      
      % Documentation: 
      % https://docs.microsoft.com/en-us/office/vba/api/excel.range(object)
      % https://docs.microsoft.com/en-us/office/vba/api/excel.xlhalign
      % https://docs.microsoft.com/en-us/office/vba/api/excel.xllinestyle
      % https://docs.microsoft.com/en-us/office/vba/api/excel.xlborderweight
      % https://docs.microsoft.com/en-us/office/vba/api/excel.borders.color
      % https://docs.microsoft.com/en-us/office/vba/api/excel.xlorientation
      % https://docs.microsoft.com/en-us/office/vba/api/excel.cellformat.interior
      
      if(self.opened)
        if(isnumeric(sheet) && sheet > 0 && sheet <= self.SheetNum)
          if(length(col) == length(row))
            rgb = @(r, g, b) r*1+g*256+b*256^2;
            p = inputParser;
            checkLogical =  @(x) islogical(x) || (isnumeric(x) && ...
              length(x)==1 && (x == 1 || x == 0)) || (ischar(x) && ...
              (strcmpi(x, 'Yes') || strcmpi(x, 'No') || ...
              strcmpi(x, 'True') || strcmpi(x, 'False')));
            
            checkEnum = @(enum, x) (isnumeric(x) && length(x)==1 && ...
              any(ismember(cell2mat(struct2cell(enum)), x))) || ...
              (ischar(x) && any(ismember(fieldnames(enum), x)));
            
            checkBorder = @(x) iscell(x) && length(x) == 3 && ...
              (isempty(x{1}) || checkEnum(self.XlBorderWeight, x{1})) ...
              && (isempty(x{2}) || isnumeric(x{2}) && length(x{2})==3) ...
              && (isempty(x{3}) || checkEnum(self.XlLineStyle, x{3}));
            
            addParameter(p, 'CellInteriorColor', [], ...
              @(x) isnumeric(x) && length(x)==3);
            addParameter(p, 'CharBold', [], checkLogical);
            addParameter(p, 'CharItalic', [], checkLogical);
            addParameter(p, 'CharUnderline', [], checkLogical);
            addParameter(p, 'CharStrikethrough', [], checkLogical);
            addParameter(p, 'CharFontName', [], @(x) ischar(x));
            addParameter(p, 'CharFontSize', [], ...
               @(x) isnumeric(x) && length(x)==1);
            addParameter(p, 'CharColor', [], ...
              @(x) isnumeric(x) && length(x)==3);
            addParameter(p, 'HorizontalAlignment', [], ...
              @(x) checkEnum(self.XlHAlign, x));
            addParameter(p, 'VerticalAlignment', [], ...
              @(x) checkEnum(self.XlVAlign, x));
            addParameter(p, 'OrientationAngle', [], ...
              @(x) isnumeric(x) && length(x)==1);
            addParameter(p, 'Orientation', [], ...
              @(x) checkEnum(self.XlOrientation, x));
            addParameter(p, 'Borders', [], checkBorder);
            addParameter(p, 'TopBorder', [], checkBorder);
            addParameter(p, 'BottomBorder', [], checkBorder);
            addParameter(p, 'LeftBorder', [], checkBorder);
            addParameter(p, 'RightBorder', [], checkBorder);
            p.parse(varargin{:})
          
            Sheet = self.Sheets.Item(sheet);
            coli = self.getColumnLetter(col);
            for i = 1:length(row)
              if(isnumeric(col) && isnumeric(row) && ...
                  all(col > 0) && all(row > 0))
                Cell = Sheet.Range([coli{i}, num2str(row(i))]);
                
                if(~isempty(p.Results.CellInteriorColor)) 
                  c = p.Results.CellInteriorColor;
                  Cell.Interior.Color = rgb(c(1), c(2), c(3));
                end
                if(~isempty(p.Results.CharBold)) 
                  Cell.Font.Bold = self.getLogical(p.Results.CharBold);
                end
                if(~isempty(p.Results.CharItalic)) 
                  Cell.Font.Italic = self.getLogical(p.Results.CharItalic);
                end
                if(~isempty(p.Results.CharUnderline)) 
                  Cell.Font.Underline = ...
                    self.getLogical(p.Results.CharUnderline);
                end
                if(~isempty(p.Results.CharStrikethrough)) 
                  Cell.Font.Strikethrough = ...
                    self.getLogical(p.Results.CharStrikethrough);
                end
                if(~isempty(p.Results.CharFontName)) 
                  Cell.Font.Name = p.Results.CharFontName;
                end
                if(~isempty(p.Results.CharFontSize)) 
                  Cell.Font.Size = p.Results.CharFontSize;
                end
                if(~isempty(p.Results.CharColor)) 
                  c = p.Results.CharColor;
                  Cell.Font.Color = rgb(c(1), c(2), c(3));
                end
                if(~isempty(p.Results.HorizontalAlignment)) 
                  Cell.HorizontalAlignment = ...
                    self.getEnumVal(self.XlHAlign, ...
                    p.Results.HorizontalAlignment);
                end
                if(~isempty(p.Results.VerticalAlignment)) 
                  Cell.VerticalAlignment = ...
                    self.getEnumVal(self.XlVAlign, ...
                    p.Results.VerticalAlignment)
                end
                if(~isempty(p.Results.OrientationAngle)) 
                  Cell.Orientation = p.Results.OrientationAngle;
                end
                if(~isempty(p.Results.Orientation)) 
                  Cell.Orientation = ...
                    self.getEnumVal(self.XlOrientation, ...
                    p.Results.Orientation)
                end
                if(~isempty(p.Results.Borders))
                  [w, crgb, ls] = self.getBorderProps(p.Results.Borders);
                  Cell.Borders.Color = crgb;
                  Cell.Borders.Weight = w;
                  Cell.Borders.LineStyle = ls;
                end
                if(~isempty(p.Results.TopBorder)) 
                  idx = self.XlBordersIndex.EdgeTop;
                  [w, crgb, ls] = ...
                    self.getBorderProps(p.Results.TopBorder);
                  Cell.Borders(idx).Color = crgb;
                  Cell.Borders(idx).Weight = w;
                  Cell.Borders(idx).LineStyle = ls;
                end
                if(~isempty(p.Results.BottomBorder)) 
                  idx = self.XlBordersIndex.EdgeBottom;
                  [w, crgb, ls] = ...
                    self.getBorderProps(p.Results.BottomBorder);
                  Cell.Borders(idx).Color = crgb;
                  Cell.Borders(idx).Weight = w;
                  Cell.Borders(idx).LineStyle = ls;
                end
                if(~isempty(p.Results.LeftBorder)) 
                  idx = self.XlBordersIndex.EdgeLeft;
                  [w, crgb, ls] = ...
                    self.getBorderProps(p.Results.LeftBorder);
                  Cell.Borders(idx).Color = crgb;
                  Cell.Borders(idx).Weight = w;
                  Cell.Borders(idx).LineStyle = ls;
                end
                if(~isempty(p.Results.RightBorder)) 
                  idx = self.XlBordersIndex.EdgeRight;
                  [w, crgb, ls] = ...
                    self.getBorderProps(p.Results.RightBorder);
                  Cell.Borders(idx).Color = crgb;
                  Cell.Borders(idx).Weight = w;
                  Cell.Borders(idx).LineStyle = ls;
                end
              else
                error('Invalid cell position!');
              end
            end
          else
            error('Different dimensions of column and row vectors!');
          end
        else
          error('Sheet does not exist!');
        end
      else
        error('No opened document!');
      end
    end
    
    function Save(self)
      if(self.opened)
        self.Document.Save;
      else
        error('No opened document!');
      end
    end

    function SaveAs(self, filePath)
      % filePath - full file path
      
      if(self.opened)
        if(exist(filePath, 'file') == 2)
          delete(filePath);
        end
        self.Document.SaveAs(filePath);
      else
        error('No opened document!');
      end
    end
    
    function coll = getColumnLetter(~, col)
      % col - column index
      % coll - column letter

      coll{length(col)} = []; n = 26; 
      for i = 1:length(col)
          coll{i} = [];
          num = col(i);
          while num > 0
             d = mod(num, n); 
             num = floor(num/n);
             if(d == 0)
                num = num-1;
                d = 26;
             end
             coll{i}(end+1) = 65+(d-1);
          end
          coll{i} = char(fliplr(coll{i}));
      end
    end
    
    function state = getLogical(~, val)
      if(islogical(val))
        state = val;
      elseif(isnumeric(val))
        state = logical(val)
      elseif(ischar(x))
        if(strcmpi(x, 'Yes') || strcmpi(x, 'True'))
          state = true;
        elseif(strcmpi(x, 'No') || strcmpi(x, 'False'))
          state = false;
        else
          error('Wrong logical type');
        end
      end
    end
    
    function out = getEnumVal(~, enum, val)
      if(isnumeric(val))
        out = val
      else
        out = enum.(val);
      end
    end
    
    function [w, crgb, ls] = getBorderProps(self, x)
      if(isempty(x{1})) 
        w = self.XlBorderWeight.Thin;
      else
        w = self.getEnumVal(self.XlBorderWeight, x{1});
      end

      if(isempty(x{2})) 
        crgb = [0, 0, 0];         
      else
        rgb = @(r, g, b) r*1+g*256+b*256^2;
        crgb = rgb(x{2}(1), x{2}(2), x{2}(3));
      end
      
      if(isempty(x{3})) 
        ls = self.XlLineStyle.Continuous;
      else
        ls = self.getEnumVal(self.XlLineStyle, x{3});
      end       
    end
  end
end
    