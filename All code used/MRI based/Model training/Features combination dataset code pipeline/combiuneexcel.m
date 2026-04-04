%% Combine Excel files using the smallest table as template (FIXED VERSION)
clear; clc;

% ===== CONFIGURATION =====
folderPath = pwd;  % Current folder (change as needed)
filePattern = '*.xlsx';  % Or '*.xls' for older files
outputFileName = 'Combined_Data_Minimal_Template.xlsx';
% =========================

% Get list of Excel files
cd(folderPath);
excelFiles = dir(filePattern);

if isempty(excelFiles)
    error('No Excel files found in folder: %s', folderPath);
end

fprintf('Found %d Excel files\n', length(excelFiles));

% Step 1: Read all files and analyze their structure
fprintf('\n📊 Step 1: Analyzing file structures...\n');
allData = cell(length(excelFiles), 1);
columnCounts = zeros(length(excelFiles), 1);
fileNames = {excelFiles.name};

for i = 1:length(excelFiles)
    filename = excelFiles(i).name;
    try
        % Read the data
        data = readtable(filename);
        allData{i} = data;
        columnCounts(i) = width(data);
        
        fprintf('  File %2d: "%s" - %d columns\n', i, filename, columnCounts(i));
        fprintf('         Columns: %s\n', strjoin(data.Properties.VariableNames, ', '));
        
    catch ME
        fprintf('  ✗ Error reading "%s": %s\n', filename, ME.message);
        allData{i} = [];  % Mark as empty on error
        columnCounts(i) = inf;  % Exclude from min calculation
    end
end

% Step 2: Find the file with the FEWEST columns (excluding errors)
[~, minIndex] = min(columnCounts);
templateFile = excelFiles(minIndex).name;
templateData = allData{minIndex};
templateColumns = templateData.Properties.VariableNames;

fprintf('\n🎯 Step 2: Using "%s" as template\n', templateFile);
fprintf('   Template has %d columns: %s\n', length(templateColumns), strjoin(templateColumns, ', '));

% Step 3: Combine all files using template structure
fprintf('\n🔄 Step 3: Combining files using template format...\n');

% Initialize combined table as empty
combinedData = table();

% Process each file and add to combined table
for i = 1:length(excelFiles)
    filename = excelFiles(i).name;
    currentData = allData{i};
    
    if isempty(currentData)
        fprintf('  ⚠ Skipping "%s" (error reading)\n', filename);
        continue;
    end
    
    % Get current file's columns
    currentColumns = currentData.Properties.VariableNames;
    
    % Check if this file has at least the template columns
    hasAllTemplateCols = all(ismember(templateColumns, currentColumns));
    
    if ~hasAllTemplateCols
        fprintf('  ⚠ "%s" is missing some template columns\n', filename);
        missingCols = setdiff(templateColumns, currentColumns);
        fprintf('     Missing: %s\n', strjoin(missingCols, ', '));
        
        % Ask user how to proceed
        response = input('     Skip this file? (y/n): ', 's');
        if lower(response) == 'y'
            fprintf('     Skipping file\n');
            continue;
        end
    end
    
    % Create a new table that matches template structure
    % Initialize as empty first, then add columns one by one
    alignedData = table();
    
    % Copy only the columns that exist in the template
    for j = 1:length(templateColumns)
        colName = templateColumns{j};
        
        if ismember(colName, currentColumns)
            % Column exists in current file, copy it
            try
                % Use the original data directly
                alignedData.(colName) = currentData.(colName);
            catch
                % If direct copy fails, try conversion
                if isnumeric(currentData.(colName))
                    alignedData.(colName) = currentData.(colName);
                else
                    % Convert to appropriate type
                    alignedData.(colName) = table2cell(currentData(:, colName));
                end
            end
        else
            % Column missing, fill with appropriate empty values based on template type
            if ~isempty(templateData)
                % Check the type from template data
                sampleValue = templateData.(colName);
                if isnumeric(sampleValue)
                    % For numeric columns, use NaN
                    alignedData.(colName) = NaN(height(currentData), 1);
                elseif islogical(sampleValue)
                    % For logical columns, use false
                    alignedData.(colName) = false(height(currentData), 1);
                elseif iscategorical(sampleValue)
                    % For categorical columns, use empty categorical
                    alignedData.(colName) = categorical(repmat({''}, height(currentData), 1));
                elseif isstring(sampleValue)
                    % For string columns, use missing string
                    alignedData.(colName) = strings(height(currentData), 1);
                    alignedData.(colName) = repmat(missing, height(currentData), 1);
                else
                    % For cell/character arrays, use empty cell
                    alignedData.(colName) = repmat({''}, height(currentData), 1);
                end
            else
                % If no template data, default to cell array of empty strings
                alignedData.(colName) = repmat({''}, height(currentData), 1);
            end
        end
    end
    
    % Add source file information
    alignedData.SourceFile = repmat({filename}, height(alignedData), 1);
    
    % Combine with main table
    if isempty(combinedData)
        combinedData = alignedData;
    else
        combinedData = [combinedData; alignedData];
    end
    
    fprintf('  ✓ Added "%s" - %d rows\n', filename, height(alignedData));
end

% Step 4: Display results and save
fprintf('\n✅ Step 4: Combination complete!\n');
fprintf('   Total rows: %d\n', height(combinedData));
fprintf('   Total columns: %d\n', width(combinedData));
fprintf('   Template columns preserved: %s\n', strjoin(templateColumns, ', '));

% Save the combined data
writetable(combinedData, outputFileName);
fprintf('\n💾 Saved to: %s\\%s\n', folderPath, outputFileName);

% Optional: Display summary statistics
if height(combinedData) > 0 && any(strcmp('SourceFile', combinedData.Properties.VariableNames))
    fprintf('\n📈 Summary by source file:\n');
    try
        summaryTable = groupsummary(combinedData, 'SourceFile');
        disp(summaryTable(:, {'SourceFile', 'GroupCount'}));
    catch
        % If groupsummary fails, do manual count
        [uniqueFiles, ~, idx] = unique(combinedData.SourceFile);
        counts = accumarray(idx, 1);
        for k = 1:length(uniqueFiles)
            fprintf('  %s: %d rows\n', uniqueFiles{k}, counts(k));
        end
    end
end

% Display preview
fprintf('\n📋 Preview of combined data (first 5 rows):\n');
disp(head(combinedData, min(5, height(combinedData))));