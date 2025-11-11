% BatchExportLensData.m
% Runs MATLABZOSConnection_data logic on every Zemax lens file in a folder
% (uses your full SurfaceData + GetNthEvenOrderTerm logic)

clc; clear;

% -------------------------------
% USER SETTINGS
% -------------------------------
lensFolder = 'C:\Users\User\OneDrive - Massachusetts Institute of Technology\Documents\MIT\Grad School\Classes\2.156\Lens Project\Prime Lenses';
outFolder  = fullfile(lensFolder, 'Exports');
if ~exist(outFolder, 'dir')
    mkdir(outFolder);
end

% -------------------------------
% CONNECT TO OPTICSTUDIO
% -------------------------------
TheApplication = MATLABZOSConnection_data();  % your full exporter also returns connection
if ischar(TheApplication) || isempty(TheApplication)
    error('Failed to connect to OpticStudio.');
end

import ZOSAPI.*;

TheSystem = TheApplication.PrimarySystem;

% -------------------------------
% GET LENS FILES
% -------------------------------
files = dir(fullfile(lensFolder, '*.zmx'));
files = [files; dir(fullfile(lensFolder, '*.zar'))];  % include .ZAR if present

fprintf('Found %d lens files in "%s".\n', numel(files), lensFolder);
fprintf('Export folder: %s\n\n', outFolder);

% -------------------------------
% MAIN LOOP
% -------------------------------
for k = 1:numel(files)
    lensPath = fullfile(files(k).folder, files(k).name);
    [~, baseName, ext] = fileparts(lensPath);
    fprintf('(%d/%d) Processing: %s\n', k, numel(files), files(k).name);

    % Load lens file
    try
        success = TheSystem.LoadFile(lensPath, false);
        if ~success
            warning('Could not load %s', lensPath);
            continue;
        end
    catch ME
        warning('Error loading %s: %s', lensPath, ME.message);
        continue;
    end

    % -------------------------------
    % COPY YOUR FULL DATA EXTRACTION LOGIC
    % -------------------------------
    try
        LDE = TheSystem.LDE;
        numSurf = LDE.NumberOfSurfaces;

        even_orders = 2:2:16;
        nEven = numel(even_orders);
        even_colnames = arrayfun(@(o) sprintf('A%d', o), even_orders, 'UniformOutput', false);

        rows = cell(numSurf, 8 + nEven);
        headers = [{'Surface','TypeName','Comment','Radius','Thickness','Material','SemiDiameter','Conic'}, even_colnames];

        for idx = 1:numSurf
            try
                surf = LDE.GetSurfaceAt(idx-1); % zero-based
            catch ME
                warning('Could not GetSurfaceAt(%d): %s', idx-1, ME.message);
                continue;
            end

            s_surface = idx-1;
            s_type = ''; s_comment = ''; s_radius = NaN; s_thickness = NaN;
            s_material = ''; s_semi = NaN; s_conic = NaN;
            s_even = nan(1, nEven);

            try, s_type = char(surf.TypeName); end
            try, s_comment = char(surf.Comment); end
            try, s_radius = double(surf.Radius); end
            try, s_thickness = double(surf.Thickness); end
            try, s_material = char(surf.Material); end
            try, s_semi = double(surf.SemiDiameter); end
            try, s_conic = double(surf.Conic); end

            surfData = [];
            try, surfData = surf.SurfaceData; catch, surfData = []; end

            if ~isempty(surfData)
                try
                    m = methods(surfData);
                    if any(strcmp(m, 'GetNthEvenOrderTerm'))
                        for n = 1:nEven
                            order = even_orders(n);
                            try
                                s_even(n) = double(surfData.GetNthEvenOrderTerm(order));
                            catch
                                try
                                    s_even(n) = double(surfData.GetNthEvenOrderTerm(n));
                                catch
                                    s_even(n) = NaN;
                                end
                            end
                        end
                    elseif any(startsWith(m, 'get_Par')) || any(contains(m, 'Par'))
                        for n = 1:nEven
                            pname = sprintf('Par%d', n);
                            try
                                pcell = surfData.(pname);
                                try
                                    s_even(n) = double(pcell.DoubleValue);
                                catch
                                    try
                                        s_even(n) = double(pcell.Value);
                                    catch
                                        try
                                            s_even(n) = str2double(char(pcell.StringValue));
                                        catch
                                            s_even(n) = NaN;
                                        end
                                    end
                                end
                            catch
                                s_even(n) = NaN;
                            end
                        end
                    elseif any(strcmp(m, 'GetCoefficient')) || any(strcmp(m, 'GetNthTerm')) || any(strcmp(m, 'GetTerm'))
                        for n = 1:nEven
                            try
                                s_even(n) = double(surfData.GetCoefficient(n));
                            catch
                                try
                                    s_even(n) = double(surfData.GetCoefficient(even_orders(n)));
                                catch
                                    s_even(n) = NaN;
                                end
                            end
                        end
                    end
                catch
                end
            end

            rowcell = cell(1, 8 + nEven);
            rowcell(1:8) = {s_surface, s_type, s_comment, s_radius, s_thickness, s_material, s_semi, s_conic};
            for n = 1:nEven
                rowcell{8 + n} = s_even(n);
            end
            rows(idx, :) = rowcell;
        end

        T = cell2table(rows, 'VariableNames', headers);

        % Write CSV with lens name
        outCSV = fullfile(outFolder, [baseName '_LensData.csv']);
        writetable(T, outCSV);
        fprintf('  → Exported %d surfaces to %s\n', height(T), outCSV);

    catch ME
        warning('Error extracting data from %s: %s', files(k).name, ME.message);
    end

    % Close file before next one
    try
        TheSystem.Close(false);
    catch
        warning('Could not close %s cleanly.', files(k).name);
    end
end

fprintf('\nAll files processed.\n');

% -------------------------------
% CLEAN UP
% -------------------------------
try
    TheApplication.CloseApplication();
catch
    warning('Could not close OpticStudio instance.');
end

fprintf('OpticStudio connection closed.\n');
