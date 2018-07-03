% Computes indifference points, moves files, writes to master sheet,
% calculates AUC and calculates k/hyperbolic
% Christine L Kuryla
% Last Updated: 20180703

% T:\Behavioral Experiments and Data\Delay Discounting\Money Discounting
% Task\money_dd_auto_20180702.m

% ****************************************************************
% Money DD - get raw data from spreadsheet
% ****************************************************************

row = input('Row Number: ');

ddRoot = fullfile('T:\Behavioral Experiments and Data\Delay Discounting');
moneyTaskFolder = fullfile(ddRoot, 'Money Discounting Task');
moneyDdDataXlsxPath = fullfile(ddRoot,'Money DD Combined Data.xlsx');

% import subject number
dcnidCell = strcat('A', int2str(row));
subjNum = xlsread(moneyDdDataXlsxPath,'Raw Data',dcnidCell);
subj = num2str(subjNum); % if there is an error, check that the cell in excel is in number format

% import date of study and convert into desired format (mmddyy)
dosRange = strcat('V', int2str(row),':X', int2str(row));
dosArray = xlsread(moneyDdDataXlsxPath,'Raw Data',dosRange);

month = num2str(dosArray(1));
day = num2str(dosArray(2));
yy = extractAfter(num2str(dosArray(3)),2);

% pad month with zero if necessary
    if length(month)==1
        mm = strcat('0',month);
    elseif length(month)==2
        mm = month;
    else
        fprintf('There is an error with the date of study!');
    end
    
% pad day with zero if necessary
    if length(day)==1
            dd = strcat('0',day);
        elseif length(day)==2
            dd = day;
    end

% Subject date of study in necessary format
subj_DOS = strcat(subj,'_',mm,dd,yy);

% ****************************************************************
% Money DD - create subject specific folder, move files
% ****************************************************************

% Create subject folder in data directory
subjFolder = fullfile(moneyTaskFolder, 'Data', subj_DOS);
fprintf('Creating folder %s... \n',subj_DOS);
mkdir(subjFolder);

% Copy the 4 excel files to the new folder
% then cd into the new folder
fprintf('Copying the 4 excel files to the new folder %s... \n',subj_DOS);
copyfile(fullfile(moneyTaskFolder, strcat(subj,'*')), subjFolder);
cd(subjFolder)

% define paths for the autogenerated csv files
indiffCsvRawName = ls(strcat(subj,'*Indiff.csv'));
indiffCsvRawPath = fullfile(subjFolder, indiffCsvRawName);
indiffTableNameNoExt = extractBefore(indiffCsvRawName,length(indiffCsvRawName)-3);

% ****************************************************************
% Use excel macro to calculate indifference points, for method see:
% https://www.mathworks.com/matlabcentral/answers/100938-how-can-i-run-an-excel-macro-from-matlab
% ****************************************************************

% Create object.
ExcelApp = actxserver('Excel.Application');

% Show window (optional).
ExcelApp.Visible = 1;

% define path to excel macro 
indiffMacroXlsPath = fullfile(moneyTaskFolder,'Data','indiff macro.xls');

% Open file located in the current folder.
cd(fullfile(moneyTaskFolder,'Data'));
ExcelApp.Workbooks.Open(indiffMacroXlsPath);
% ExcelApp.Workbooks.Open(fullfile(pwd,'\myFile.xls')); %%%%%%%%%%%%

% Run Macro1, defined in "ThisWorkBook" with one parameter. 
% A return value cannot be retrieved
% ExcelApp.Run('ThisWorkBook.Macro1', parameter1);
% Run Macro2, defined in "Sheet1" with two parameters. 
%A return value cannot be retrieved.
% ExcelApp.Run('Sheet1.Macro2', parameter1, parameter2);
% Run Macro3, defined in the module "Module1" 
% with no parameters and a return value.
% retVal = ExcelApp.Run('Macro3');

indiffRaw = fullfile(moneyTaskFolder, strcat(subj,'*Indiff.csv'));
ExcelApp.Run('indiff', indiffRaw);

% Quit application and release object.
ExcelApp.Quit;
ExcelApp.release;

% ****************************************************************
% Use results from indiff macro
% (copy indiff points to master money dd sheet)
% ****************************************************************

% Macro resulted in the creation of a new xlsx file, so define the path
indiffCalcXlsxPath = fullfile(subjFolder, strcat(indiffTableNameNoExt,'.xlsx'));

% Desired data is in cells B28:F28 in the new xlsx file
indiffData = xlsread(indiffCalcXlsxPath, indiffTableNameNoExt, 'B28:F28');

% want to input indiff points into G374:K374 on master money sheet
rangeOut = strcat('G', int2str(row), ':K', int2str(row));

% write data to sheet
fprintf('Writing %s indiff data to %s \n',subj_DOS,rangeOut);
xlswrite(moneyDdDataXlsxPath, indiffData, 'Raw Data', rangeOut);

fprintf("%4.2f ",indiffData);
fprintf("\n Finished %s indiff calc!\n",subjDOS); 


% ****************************************************************
% Money DD - calculate AUC 
% ****************************************************************

moneyD1 = indiffData(2);
moneyD7 = indiffData(3);
moneyD30 = indiffData(4);
moneyD90 = indiffData(5);

% Algorithm from excel spreadsheet ddRoot\Game DD Discounting Calculator.xlsx % gameDdCalc = fullfile(ddRoot, 'Game DD Discounting Calculator.xlsx');
% We are approximating AUC (area under the curve) by summing the constituent 
% trapezoids, then noting fraction of max AUC

% Parameters:
A = 10; % from sheet
maxDelay = 90; % (max of 1,7,30,90)

% Area of trapezoid = .5*(b1+b2)*h
areaTrap1 = .5*moneyD1*1;
areaTrap2 = .5*(moneyD1+moneyD7)*(7-1);
areaTrap3 = .5*(moneyD7+moneyD30)*(30-7);
areaTrap4 = .5*(moneyD30+moneyD90)*(90-30);

% Fractional AUC --- this was the goal
MoneyDD_AUC = (areaTrap1+areaTrap2+areaTrap3+areaTrap4) / (A*maxDelay) ;

% Write to master money data sheet
aucCell = strcat('Q', int2str(row)); % AUC is column Q

xlswrite(moneyDdDataXlsxPath, MoneyDD_AUC, 'Raw Data', aucCell);

fprintf('AUC is %d \n',MoneyDD_AUC);

% ****************************************************************
% Money DD - calculate k and hyperbolic
% ****************************************************************

% set parameters

% k (constant from sheet)
k = 0.0195253849422891;