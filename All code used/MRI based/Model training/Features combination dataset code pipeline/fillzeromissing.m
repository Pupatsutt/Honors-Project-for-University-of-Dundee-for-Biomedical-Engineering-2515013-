%% Ultra-Simple MATLAB Script: Fill Missing Values with 0
clear; clc;

%% Read the data
data = readtable('Combined_Data_Minimal_Template.xlsx', 'VariableNamingRule', 'preserve');
fprintf('Original data: %d rows, %d columns\n', height(data), width(data));

%% Loop through each column and fill missing values
filledCount = 0;

for i = 1:width(data)
    colName = data.Properties.VariableNames{i};
    
    % Check if column is numeric
    if isnumeric(data{:, i})
        nanIdx = isnan(data{:, i});
        if any(nanIdx)
            data{nanIdx, i} = 0;
            filledCount = filledCount + sum(nanIdx);
            fprintf('Filled %d missing in numeric column: %s\n', sum(nanIdx), colName);
        end
    end
    
    % Check if column is cell array (text)
    if iscell(data{:, i})
        emptyIdx = cellfun(@isempty, data{:, i});
        if any(emptyIdx)
            data{emptyIdx, i} = {'0'};
            filledCount = filledCount + sum(emptyIdx);
            fprintf('Filled %d empty cells in text column: %s\n', sum(emptyIdx), colName);
        end
    end
end

%% Save the result
writetable(data, 'Combined_Data_Filled_With_Zero.xlsx');
fprintf('\nDone! Filled %d missing values total.\n', filledCount);