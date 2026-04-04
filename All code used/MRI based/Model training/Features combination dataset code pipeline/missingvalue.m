%% Automated Missing Values Analysis for Combined_Data_Minimal_Template.xlsx
% This script automatically loads and analyzes the specified Excel file
% No user interaction required - runs completely automatically

%% Clear workspace and command window
clear; clc;
fprintf('========================================\n');
fprintf('MISSING VALUES ANALYSIS\n');
fprintf('========================================\n');
fprintf('Started: %s\n\n', datestr(now));

%% Define file name and path
fileName = 'Combined_Data_Minimal_Template.xlsx';

% Check if file exists in current directory
if exist(fileName, 'file') == 2
    filePath = fileName;
    fprintf('Found file in current directory: %s\n', fileName);
else
    % If not in current directory, try to find it
    fprintf('File not found in current directory. Searching...\n');
    
    % Search in current directory and subdirectories
    fileList = dir(['**/' fileName]);
    
    if ~isempty(fileList)
        filePath = fullfile(fileList(1).folder, fileList(1).name);
        fprintf('Found file at: %s\n', filePath);
    else
        % If still not found, prompt user (but then make it automatic after)
        fprintf('\nWARNING: Could not find %s\n', fileName);
        fprintf('Please place the file in the current directory or provide the full path.\n');
        
        % Try one more time with user input (but script continues automatically)
        filePath = input('Enter full path to file (or press Enter to cancel): ', 's');
        if isempty(filePath)
            error('File not found. Analysis cancelled.');
        end
    end
end

%% Load the Excel file
fprintf('\n=== LOADING DATA ===\n');
try
    % Try to read all sheets first to see what's available
    [~, sheetNames] = xlsfinfo(filePath);
    fprintf('Sheets found in file: %s\n', strjoin(sheetNames, ', '));
    
    % Read the first sheet by default (most common case)
    dataTable = readtable(filePath);
    fprintf('Successfully loaded data from first sheet\n');
    fprintf('  Rows: %d\n', height(dataTable));
    fprintf('  Columns: %d\n', width(dataTable));
    
    % Display first few rows to verify data
    fprintf('\nFirst 5 rows of data:\n');
    disp(head(dataTable, 5));
    
catch ME
    fprintf('ERROR loading file: %s\n', ME.message);
    
    % Try alternative loading methods
    try
        fprintf('\nTrying alternative loading method...\n');
        opts = detectImportOptions(filePath);
        opts.VariableNamingRule = 'preserve';
        dataTable = readtable(filePath, opts);
        fprintf('Successfully loaded with custom options\n');
    catch ME2
        error('Failed to load file: %s\n%s', ME2.message);
    end
end

%% Initialize results structure
results = struct();
results.timestamp = datestr(now);
results.fileName = filePath;
results.tableInfo = struct();

%% Get basic table information
results.tableInfo.numRows = height(dataTable);
results.tableInfo.numCols = width(dataTable);
results.tableInfo.varNames = dataTable.Properties.VariableNames;
results.tableInfo.varTypes = varfun(@class, dataTable, 'OutputFormat', 'cell');

fprintf('\n=== TABLE INFORMATION ===\n');
fprintf('Total rows: %d\n', results.tableInfo.numRows);
fprintf('Total columns: %d\n', results.tableInfo.numCols);
fprintf('\nColumn details:\n');
for i = 1:results.tableInfo.numCols
    % Get sample value
    sampleData = dataTable.(results.tableInfo.varNames{i});
    if ~isempty(sampleData) && results.tableInfo.numRows > 0
        sampleVal = sampleData(1);
        if iscell(sampleVal)
            sampleStr = char(sampleVal);
        elseif isnumeric(sampleVal) || islogical(sampleVal)
            sampleStr = num2str(sampleVal);
        elseif isstring(sampleVal)
            sampleStr = char(sampleVal);
        else
            sampleStr = class(sampleVal);
        end
    else
        sampleStr = 'N/A';
    end
    fprintf('  %2d. %-25s (%-10s) e.g.: %s\n', ...
        i, results.tableInfo.varNames{i}, results.tableInfo.varTypes{i}, sampleStr);
end

%% Analyze missing values
fprintf('\n=== MISSING VALUES ANALYSIS ===\n');

% Initialize arrays for statistics
missingCounts = zeros(1, results.tableInfo.numCols);
missingPercent = zeros(1, results.tableInfo.numCols);
dataTypes = cell(1, results.tableInfo.numCols);
missingPatterns = cell(1, results.tableInfo.numCols);

% Create missing matrix
missingMatrix = false(results.tableInfo.numRows, results.tableInfo.numCols);

% Header for table
fprintf('\n%-3s %-25s %-12s %-10s %-10s %s\n', ...
    '#', 'Column Name', 'Data Type', 'Missing', 'Missing %', 'Missing Values');
fprintf('%s\n', repmat('-', 1, 90));

for i = 1:results.tableInfo.numCols
    varName = results.tableInfo.varNames{i};
    varData = dataTable.(varName);
    dataTypes{i} = class(varData);
    
    % Detect missing values based on data type
    if isnumeric(varData)
        % Numeric: check NaN and Inf
        missingIdx = isnan(varData) | isinf(varData);
        missingVals = {};
        if any(isnan(varData))
            missingVals{end+1} = 'NaN';
        end
        if any(isinf(varData))
            missingVals{end+1} = 'Inf';
        end
        missingMatrix(:,i) = missingIdx;
        
    elseif iscategorical(varData)
        % Categorical: check undefined
        missingIdx = isundefined(varData);
        missingVals = {'<undefined>'};
        missingMatrix(:,i) = missingIdx;
        
    elseif isstring(varData)
        % String: check missing and empty
        missingIdx = ismissing(varData) | (varData == "") | (varData == "NaN") | (varData == "NA") | (varData == "NULL");
        missingVals = {};
        if any(ismissing(varData))
            missingVals{end+1} = '<missing>';
        end
        if any(varData == "")
            missingVals{end+1} = 'empty';
        end
        if any(varData == "NaN")
            missingVals{end+1} = '"NaN"';
        end
        if any(varData == "NA")
            missingVals{end+1} = '"NA"';
        end
        if any(varData == "NULL")
            missingVals{end+1} = '"NULL"';
        end
        missingMatrix(:,i) = missingIdx;
        
    elseif iscell(varData)
        % Cell array: comprehensive check
        missingIdx = false(size(varData));
        missingVals = {};
        
        % Check for various missing representations
        for j = 1:length(varData)
            val = varData{j};
            if isempty(val)
                missingIdx(j) = true;
                missingVals{end+1} = 'empty';
            elseif ischar(val) && (strcmpi(val, '') || strcmpi(val, 'nan') || strcmpi(val, 'na') || strcmpi(val, 'null'))
                missingIdx(j) = true;
                missingVals{end+1} = upper(val);
            elseif isstring(val) && (ismissing(val) || val == "" || val == "NaN" || val == "NA")
                missingIdx(j) = true;
                missingVals{end+1} = char(val);
            elseif isnumeric(val) && isnan(val)
                missingIdx(j) = true;
                missingVals{end+1} = 'NaN';
            end
        end
        missingVals = unique(missingVals);
        missingMatrix(:,i) = missingIdx;
        
    elseif islogical(varData)
        % Logical: usually can't be missing
        missingIdx = false(size(varData));
        missingVals = {'none'};
        missingMatrix(:,i) = missingIdx;
        
    elseif isdatetime(varData)
        % Datetime: check NaT
        missingIdx = isnat(varData);
        missingVals = {'NaT'};
        missingMatrix(:,i) = missingIdx;
        
    else
        % Other types
        try
            missingIdx = ismissing(varData);
            missingVals = {'unknown'};
        catch
            missingIdx = false(size(varData));
            missingVals = {'not detectable'};
        end
        missingMatrix(:,i) = missingIdx;
    end
    
    % Store statistics
    missingCounts(i) = sum(missingIdx);
    missingPercent(i) = (missingCounts(i) / results.tableInfo.numRows) * 100;
    missingPatterns{i} = strjoin(unique(missingVals), ', ');
    
    % Display row
    fprintf('%-3d %-25s %-12s %-10d %-9.2f%% %s\n', ...
        i, ...
        strtrunc(varName, 25), ...
        strtrunc(dataTypes{i}, 12), ...
        missingCounts(i), ...
        missingPercent(i), ...
        strtrunc(missingPatterns{i}, 30));
end

%% Store column statistics
results.columnStats = table(...
    (1:results.tableInfo.numCols)', ...
    results.tableInfo.varNames', ...
    dataTypes', ...
    missingCounts', ...
    missingPercent', ...
    missingPatterns', ...
    'VariableNames', {'Index', 'Variable', 'DataType', 'MissingCount', 'MissingPercent', 'MissingTypes'});

%% Overall summary statistics
fprintf('\n=== OVERALL SUMMARY ===\n');

% Total missing values
results.summary.totalMissing = sum(missingCounts);
results.summary.totalCells = results.tableInfo.numRows * results.tableInfo.numCols;
results.summary.overallMissingPercent = (results.summary.totalMissing / results.summary.totalCells) * 100;

fprintf('Total missing values: %d out of %d cells (%.2f%%)\n', ...
    results.summary.totalMissing, results.summary.totalCells, results.summary.overallMissingPercent);

% Columns with missing data
colsWithMissing = missingCounts > 0;
results.summary.numColsWithMissing = sum(colsWithMissing);
results.summary.colsWithMissingNames = results.tableInfo.varNames(colsWithMissing);
results.summary.colsWithMissingPercent = missingPercent(colsWithMissing);

fprintf('Columns with missing data: %d out of %d (%.2f%%)\n', ...
    results.summary.numColsWithMissing, results.tableInfo.numCols, ...
    (results.summary.numColsWithMissing/results.tableInfo.numCols)*100);

if results.summary.numColsWithMissing > 0
    fprintf('  Columns with highest missing %:\n');
    [sortedPct, sortIdx] = sort(results.summary.colsWithMissingPercent, 'descend');
    for j = 1:min(5, length(sortedPct))
        colName = results.summary.colsWithMissingNames{sortIdx(j)};
        fprintf('    %d. %s: %.2f%%\n', j, colName, sortedPct(j));
    end
end

% Rows with missing data
rowsWithMissing = any(missingMatrix, 2);
results.summary.numRowsWithMissing = sum(rowsWithMissing);
results.summary.rowsWithMissingPercent = (results.summary.numRowsWithMissing / results.tableInfo.numRows) * 100;

fprintf('\nRows with missing data: %d out of %d (%.2f%%)\n', ...
    results.summary.numRowsWithMissing, results.tableInfo.numRows, ...
    results.summary.rowsWithMissingPercent);

% Calculate missingness distribution
missingPerRow = sum(missingMatrix, 2);
results.summary.rowMissingDistribution = histcounts(missingPerRow, 0:max(missingPerRow)+1);

fprintf('\nMissing data distribution by row:\n');
fprintf('  %-20s: %d rows\n', 'Rows with 0 missing', sum(missingPerRow == 0));
for k = 1:max(missingPerRow)
    count = sum(missingPerRow == k);
    if count > 0
        fprintf('  %-20s: %d rows\n', sprintf('Rows with %d missing', k), count);
    end
end

%% Generate detailed report for specific columns of interest
fprintf('\n=== DETAILED ANALYSIS FOR COLUMNS WITH >20%% MISSING ===\n');
highMissingCols = find(missingPercent > 20);

if ~isempty(highMissingCols)
    for i = 1:length(highMissingCols)
        colIdx = highMissingCols(i);
        colName = results.tableInfo.varNames{colIdx};
        colData = dataTable.(colName);
        
        fprintf('\nColumn: %s (%.2f%% missing)\n', colName, missingPercent(colIdx));
        
        % Show sample of missing values
        missingRows = find(missingMatrix(:, colIdx));
        if ~isempty(missingRows)
            fprintf('  First 5 missing value rows: %s\n', ...
                strjoin(cellstr(num2str(missingRows(1:min(5, end))')), ', '));
            
            % Show what the missing values look like
            fprintf('  Missing value examples:\n');
            for j = 1:min(3, length(missingRows))
                rowNum = missingRows(j);
                val = colData(rowNum);
                if iscell(val)
                    valStr = char(val);
                elseif isnumeric(val)
                    valStr = num2str(val);
                elseif isstring(val)
                    valStr = char(val);
                else
                    valStr = class(val);
                end
                fprintf('    Row %d: "%s"\n', rowNum, valStr);
            end
        end
    end
else
    fprintf('No columns with >20%% missing values found.\n');
end

%% Generate visualizations
fprintf('\n=== GENERATING VISUALIZATIONS ===\n');

try
    % Create figure
    fig = figure('Name', 'Missing Values Analysis - Combined_Data_Minimal_Template', ...
        'Position', [50, 50, 1400, 900], 'Visible', 'on');
    
    % 1. Missing values heatmap
    subplot(2, 3, [1, 2]);
    imagesc(missingMatrix');
    colormap(gca, [1 1 1; 0.8 0 0]);
    colorbar('Ticks', [0, 1], 'TickLabels', {'Present', 'Missing'});
    title('Missing Values Heatmap');
    set(gca, 'YTick', 1:results.tableInfo.numCols, 'YTickLabel', results.tableInfo.varNames);
    xlabel('Row Number');
    ylabel('Variables');
    
    % 2. Missing percentage bar chart
    subplot(2, 3, 3);
    [sortedPercent, sortIdx] = sort(missingPercent, 'descend');
    barh(sortedPercent, 'FaceColor', [0.3 0.6 0.9]);
    title('Missing % by Column');
    xlabel('Missing Percentage (%)');
    ylabel('Columns');
    set(gca, 'YTick', 1:results.tableInfo.numCols, 'YTickLabel', results.tableInfo.varNames(sortIdx));
    xlim([0, 100]);
    grid on;
    
    % 3. Missing count bar chart
    subplot(2, 3, 4);
    bar(missingCounts, 'FaceColor', [0.9 0.4 0.2]);
    title('Missing Count by Column');
    xlabel('Column Index');
    ylabel('Missing Count');
    set(gca, 'XTick', 1:results.tableInfo.numCols, 'XTickLabel', results.tableInfo.varNames, ...
        'XTickLabelRotation', 45);
    grid on;
    
    % 4. Distribution of missing values per row
    subplot(2, 3, 5);
    histogram(missingPerRow, 'FaceColor', [0.2 0.7 0.3], 'BinMethod', 'integers');
    title('Missing Values per Row');
    xlabel('Number of Missing Values');
    ylabel('Number of Rows');
    grid on;
    
    % 5. Data types pie chart
    subplot(2, 3, 6);
    [uniqueTypes, ~, typeIdx] = unique(dataTypes);
    typeCounts = histcounts(typeIdx, 0.5:length(uniqueTypes)+0.5);
    pie(typeCounts, uniqueTypes);
    title('Data Type Distribution');
    
    % Overall title
    sgtitle(sprintf('Missing Values Analysis - %s\nTotal: %.2f%% missing', ...
        fileName, results.summary.overallMissingPercent));
    
    % Save figure
    saveas(fig, 'Missing_Values_Analysis_Combined_Data.png');
    fprintf('  Visualization saved to: Missing_Values_Analysis_Combined_Data.png\n');
    
catch ME
    fprintf('  Warning: Could not generate visualizations: %s\n', ME.message);
end

%% Save results to files
fprintf('\n=== SAVING RESULTS ===\n');

% Save column statistics to CSV
csvFile = 'Missing_Values_Summary_Combined_Data.csv';
writetable(results.columnStats, csvFile);
fprintf('  Column statistics saved to: %s\n', csvFile);

% Save detailed text report
reportFile = 'Missing_Values_Report_Combined_Data.txt';
fid = fopen(reportFile, 'w');

fprintf(fid, '=');
fprintf(fid, 'MISSING VALUES ANALYSIS REPORT\n');
fprintf(fid, '=');
fprintf(fid, '\n');
fprintf(fid, 'File: %s\n', fileName);
fprintf(fid, 'Generated: %s\n\n', results.timestamp);

fprintf(fid, 'TABLE INFORMATION\n');
fprintf(fid, '-----------------\n');
fprintf(fid, '  Total Rows: %d\n', results.tableInfo.numRows);
fprintf(fid, '  Total Columns: %d\n\n', results.tableInfo.numCols);

fprintf(fid, 'OVERALL SUMMARY\n');
fprintf(fid, '---------------\n');
fprintf(fid, '  Total missing cells: %d (%.2f%%)\n', results.summary.totalMissing, results.summary.overallMissingPercent);
fprintf(fid, '  Columns with missing: %d (%.2f%%)\n', results.summary.numColsWithMissing, ...
    (results.summary.numColsWithMissing/results.tableInfo.numCols)*100);
fprintf(fid, '  Rows with missing: %d (%.2f%%)\n\n', results.summary.numRowsWithMissing, results.summary.rowsWithMissingPercent);

fprintf(fid, 'COLUMN DETAILS\n');
fprintf(fid, '--------------\n');
for i = 1:height(results.columnStats)
    fprintf(fid, '  %2d. %-25s (%-12s): %5d missing (%6.2f%%) - %s\n', ...
        results.columnStats.Index(i), ...
        results.columnStats.Variable{i}, ...
        results.columnStats.DataType{i}, ...
        results.columnStats.MissingCount(i), ...
        results.columnStats.MissingPercent(i), ...
        results.columnStats.MissingTypes{i});
end

fprintf(fid, '\nCOLUMNS WITH HIGH MISSING (>20%%)\n');
fprintf(fid, '--------------------------------\n');
highMissing = results.columnStats(results.columnStats.MissingPercent > 20, :);
if ~isempty(highMissing)
    for i = 1:height(highMissing)
        fprintf(fid, '  %s: %.2f%% missing\n', highMissing.Variable{i}, highMissing.MissingPercent(i));
    end
else
    fprintf(fid, '  No columns with >20%% missing values\n');
end

fclose(fid);
fprintf('  Detailed report saved to: %s\n', reportFile);

% Save workspace with all results
matFile = 'Missing_Values_Results_Combined_Data.mat';
save(matFile, 'results', 'dataTable', 'missingMatrix', 'missingCounts', 'missingPercent');
fprintf('  Complete results saved to: %s\n', matFile);

%% Final summary
fprintf('\n========================================\n');
fprintf('ANALYSIS COMPLETE\n');
fprintf('========================================\n');
fprintf('Files generated:\n');
fprintf('  1. Missing_Values_Analysis_Combined_Data.png (visualization)\n');
fprintf('  2. Missing_Values_Summary_Combined_Data.csv (column statistics)\n');
fprintf('  3. Missing_Values_Report_Combined_Data.txt (detailed report)\n');
fprintf('  4. Missing_Values_Results_Combined_Data.mat (complete data)\n');
fprintf('\nFinished: %s\n', datestr(now));

%% Helper function to truncate strings
function truncated = strtrunc(str, maxLen)
    if length(str) > maxLen
        truncated = [str(1:maxLen-3) '...'];
    else
        truncated = str;
    end
end