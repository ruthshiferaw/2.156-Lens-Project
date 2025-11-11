% BatchExportAnalyses.m
% Runs MATLABZOSConnection_analysis logic for every Zemax lens file
% in a given folder and exports analysis results.

clc; clear;

% -------------------------------
% USER SETTINGS
% -------------------------------
lensFolder = 'C:\Users\User\OneDrive - Massachusetts Institute of Technology\Documents\MIT\Grad School\Classes\2.156\Lens Project\Prime Lenses';
outRoot = fullfile(lensFolder, 'AnalysisExports');
if ~exist(outRoot, 'dir')
    mkdir(outRoot);
end

% -------------------------------
% CONNECT TO OPTICSTUDIO
% -------------------------------
TheApplication = MATLABZOSConnection_analysis(); % returns connection if successful
if ischar(TheApplication) || isempty(TheApplication)
    error('Failed to connect to OpticStudio.');
end

import ZOSAPI.*;

TheSystem = TheApplication.PrimarySystem;

% -------------------------------
% GET LENS FILES
% -------------------------------
files = dir(fullfile(lensFolder, '*.zmx'));
files = [files; dir(fullfile(lensFolder, '*.zar'))];

fprintf('Found %d lens files in "%s".\n\n', numel(files), lensFolder);

% -------------------------------
% MAIN LOOP
% -------------------------------
for k = 1:numel(files)
    lensPath = fullfile(files(k).folder, files(k).name);
    [~, baseName, ~] = fileparts(lensPath);
    fprintf('(%d/%d) Running analyses for: %s\n', k, numel(files), files(k).name);

    % Create an export subfolder for each lens
    outDir = fullfile(outRoot, baseName);
    if ~exist(outDir, 'dir')
        mkdir(outDir);
    end

    % Load the lens file
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

    % -------------------------------------------------------
    % COPY YOUR ANALYSIS EXPORT LOGIC HERE
    % (adapted to save each lens in its own outDir)
    % -------------------------------------------------------
    try
        analyses = TheSystem.Analyses;
        sysName = char(TheSystem.SystemFile);
        [~, baseName, ~] = fileparts(sysName);

        % --- Field Curvature ---
        try
            fc = analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.FieldCurvatureAndDistortion);
            fc.ApplyAndWaitForCompletion();
            fcFile = fullfile(outDir, [baseName '_FieldCurvature.txt']);
            fcStream = System.IO.StreamWriter(fcFile);
            fc.GetResults().WriteToStream(fcStream);
            fcStream.Close();
            fc.Close();
        catch ME
            warning('Field curvature export failed for %s: %s', baseName, ME.message);
        end
        % 
        % % --- Vignetting ---
        % try
        %     vig = analyses.New_RelativeIllumination();
        %     vig.ApplyAndWaitForCompletion();
        %     vigFile = fullfile(outDir, [baseName '_Vignetting.txt']);
        %     vigStream = System.IO.StreamWriter(vigFile);
        %     vig.GetResults().WriteToStream(vigStream);
        %     vigStream.Close();
        %     vig.Close();
        % catch ME
        %     warning('Vignetting export failed for %s: %s', baseName, ME.message);
        % end
        %
        % % --- RMS v Field ---
        % try
        %     RMS = analyses.New_RMSField();
        %     RMS.ApplyAndWaitForCompletion();
        %     RMSFile = fullfile(outDir, [baseName '_RMSvField.txt']);
        %     RMSStream = System.IO.StreamWriter(RMSFile);
        %     RMS.GetResults().WriteToStream(RMSStream);
        %     RMSStream.Close();
        %     RMS.Close();
        % catch ME
        %     warning('Vignetting export failed for %s: %s', baseName, ME.message);
        % end
        % 
        % % --- Spot Diagram (3 fields) ---
        % try
        %     spot = analyses.New_StandardSpot();
        %     numFields = 3; % adjust if needed
        %     for f = 1:numFields
        %         spot.Field.SetField(f);
        %         spot.ApplyAndWaitForCompletion();
        %         spotFile = fullfile(outDir, sprintf('%s_SpotDiagram_Field%d.txt', baseName, f));
        %         spotStream = System.IO.StreamWriter(spotFile);
        %         spot.GetResults().WriteToStream(spotStream);
        %         spotStream.Close();
        %     end
        %     spot.Close();
        % catch ME
        %     warning('Spot diagram export failed for %s: %s', baseName, ME.message);
        % end
        %
        % % --- Longitudinal Aberration ---
        % try
        %     long = analyses.New_LongitudinalAberration
        %     long.ApplyAndWaitForCompletion();
        %     longFile = fullfile(outDir, [baseName '_Longitudinal.txt']);
        %     longStream = System.IO.StreamWriter(longFile);
        %     long.GetResults().WriteToStream(longStream);
        %     longStream.Close();
        %     long.Close();;
        % catch ME
        %     warning('Longitudinal aberration export failed for %s: %s', baseName, ME.message);
        % end

        fprintf('  ✅ Analyses exported to %s\n\n', outDir);

    catch ME
        warning('Error running analyses on %s: %s', files(k).name, ME.message);
    end

    % Close system before next file
    try
        TheSystem.Close(false);
    catch
        warning('Could not close %s cleanly.', files(k).name);
    end
end

% -------------------------------
% CLEAN UP
% -------------------------------
try
    TheApplication.CloseApplication();
catch
    warning('Could not close OpticStudio instance.');
end

fprintf('All analyses complete.\n');
