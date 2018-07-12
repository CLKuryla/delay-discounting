% Automate game time task calculations for Keri's delay discounting study

% First, enter data into spreadsheet Game DD Combined Data.xlsx 
% (located in 'T:\Behavioral Experiments and Data\Delay Discounting\')
% then note row that needs to be calculated, save and close spreadsheet,
% and run this script, inputting the desired row number when prompted.

% Info from T:\Behavioral Experiments and Data\Delay Discounting
% Christine L Kuryla, last updated 20180703

% ****************************************************************
% Game DD - get raw data from spreadsheet
% ****************************************************************

% define root directory and spreadsheet location
ddRoot = fullfile('T:\Behavioral Experiments and Data\Delay Discounting');
gameDdDataXlsPath = fullfile(ddRoot,'Game DD Combined Data.xlsx');

% cd into Delay Discounting folder
cd(ddRoot)

% prompt for row number
row = input('Row Number: ');

% define strings that correspond to excel spreadsheet
dd25Cell = strcat('N', int2str(row));
dd50Cell = strcat('O', int2str(row));
dd100Cell = strcat('P', int2str(row));

% load values of interest
GameDD_D25 = xlsread(gameDdDataXlsPath,'Raw Data',dd25Cell);
GameDD_D50 = xlsread(gameDdDataXlsPath,'Raw Data',dd50Cell);
GameDD_D100 = xlsread(gameDdDataXlsPath,'Raw Data',dd100Cell);

% the following reformats is optional - if it errors, simply comment it out
dcnidCell = strcat('A', int2str(row));

% load subject ID number
subj = xlsread(gameDdDataXlsPath,'Raw Data',dcnidCell);

% aucCell = strcat('T', int2str(row));
% we will output GameDD_AUC in T312 or T306 or whatever

% print output to double check if needed
fprintf('Subject %i, in row %d, dd25, dd50, dd100 = %d,%d,%d \n',subj,row,GameDD_D25,GameDD_D50,GameDD_D100);

% ****************************************************************
% Game DD - Calculate AUC (fractional)
% ****************************************************************
% Input: GameDD_D25, GameDD_D50, GameDD_D100 (see above)
% Output: AUC (area under curve)
% Algorithm from excel spreadsheet ddRoot\Game DD Discounting Calculator.xlsx % gameDdCalc = fullfile(ddRoot, 'Game DD Discounting Calculator.xlsx');
% We are approximating AUC (area under the curve) by summing the constituent 
% trapezoids, then noting fraction of max AUC

% Parameters:
A = 60; % max dd
maxDelay = 100; % (max of 25, 50, 100)

% Area of trapezoid = .5*(b1+b2)*h
areaTrap1 = .5*GameDD_D25*25;
areaTrap2 = .5*(GameDD_D25+GameDD_D50)*(50-25);
areaTrap3 = .5*(GameDD_D50+GameDD_D100)*(100-50);

% Fractional AUC --- this was the goal
GameDD_AUC = (areaTrap1+areaTrap2+areaTrap3) / (A*maxDelay) ;

fprintf('AUC is %d \n',GameDD_AUC);

% ****************************************************************
% Game DD - output results and write to spreadsheet
% ****************************************************************



% define cell where we want to write output, write output, print update
aucCell = strcat('T', int2str(row));
xlswrite(gameDdDataXlsPath,GameDD_AUC,'Raw Data',aucCell);
fprintf('Money DD spreadsheet updated AUC %d for subject %d (in row %d) \n',GameDD_AUC,subj,row);

% close excel file
fid=fopen(gameDdDataXlsPath);
fclose(fid);

% done!
