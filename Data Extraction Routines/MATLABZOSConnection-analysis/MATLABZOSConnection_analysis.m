function [ TheApplication ] = MATLABZOSConnection_analysis( instance )

if ~exist('instance', 'var')
    instance = 0;
else
    try
        instance = int32(instance);
    catch
        instance = 0;
        warning('Invalid parameter {instance}');
    end
end

% Initialize the OpticStudio connection
TheApplication = InitConnection(instance);
if isempty(TheApplication)
    % failed to initialize a connection
    TheApplication = 'Failed to connect to OpticStudio';
else
    import ZOSAPI.*;

    % You can add custom code here.
    % Even without additional code the MATLAB command window
    % will allow you to interact directly with OpticStudio.
    

    %% Minimal analysis export script   
    TheSystem = TheApplication.PrimarySystem;
    analyses = TheSystem.Analyses;
    methods(analyses)
    TheSystem.SystemData
    
    outDir = fullfile(getenv('USERPROFILE'), 'Documents', 'ZemaxExports');
    if ~exist(outDir, 'dir')
        mkdir(outDir);
    end
    
    % sysName = string(TheSystem.SystemFile);
    % [~, baseName, ~] = fileparts(sysName);  % removes path and extension
    baseName = 'MyLensSystem'; % safer fallback if system name unavailable
    sysName = char(TheSystem.SystemFile);
    [~, baseName, ~] = fileparts(sysName);  % remove path and extension
    
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
        warning('Field curvature export failed: %s', ME.message);
    end
    
    % --- Vignetting ---
    try
        vig = analyses.New_RelativeIllumination();
        vig.ApplyAndWaitForCompletion();
        vigFile = fullfile(outDir, [baseName '_Vignetting.txt']);
        vigStream = System.IO.StreamWriter(vigFile);
        vig.GetResults().WriteToStream(vigStream);
        vigStream.Close();
        vig.Close();
    catch ME
        warning('Vignetting export failed: %s', ME.message);
    end
    
    % --- Spot Diagram ---
    try
        spot = analyses.New_StandardSpot();
        spot.ApplyAndWaitForCompletion();

        spotFile = fullfile(outDir, [baseName '_SpotDiagram.txt']);
        spotStream = System.IO.StreamWriter(spotFile);
        spot.GetResults().WriteToStream(spotStream);
        spotStream.Close();
        spot.Close();

        
        % spot = analyses.New_StandardSpot();
        % 
        % numFields = TheSystem.SystemData.Fields.NumberOfFields;  % set manually, or parse from Field Data table in the file
        % for f = 1:numFields
        %     spot.Field.SetField(f);
        %     spot.ApplyAndWaitForCompletion();
        % 
        %     % save each field separately
        %     spotFile = fullfile(outDir, sprintf('%s_SpotDiagram_Field%d.txt', baseName, f));
        %     spotStream = System.IO.StreamWriter(spotFile);
        %     spot.GetResults().WriteToStream(spotStream);
        %     spotStream.Close();
        % end
        % spot.Close();

    catch ME
        warning('Spot diagram export failed: %s', ME.message);
    end
    
    disp(['✅ Analyses exported to ' outDir]);

end
end

function app = InitConnection(instance)

import System.Reflection.*;

% Find the installed version of OpticStudio.
zemaxData = winqueryreg('HKEY_CURRENT_USER', 'Software\Zemax', 'ZemaxRoot');
NetHelper = strcat(zemaxData, '\ZOS-API\Libraries\ZOSAPI_NetHelper.dll');
% Note -- uncomment the following line to use a custom NetHelper path
% NetHelper = 'C:\Users\User\OneDrive\Documents\Zemax\ZOS-API\Libraries\ZOSAPI_NetHelper.dll';
% This is the path to OpticStudio
NET.addAssembly(NetHelper);

success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
% Note -- uncomment the following line to use a custom initialization path
% success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize('C:\Program Files\OpticStudio\');
if success == 1
    LogMessage(strcat('Found OpticStudio at: ', char(ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory())));
else
    app = [];
    return;
end

% Now load the ZOS-API assemblies
NET.addAssembly(AssemblyName('ZOSAPI_Interfaces'));
NET.addAssembly(AssemblyName('ZOSAPI'));

% Create the initial connection class
TheConnection = ZOSAPI.ZOSAPI_Connection();

% Attempt to create a Standalone connection

% NOTE - if this fails with a message like 'Unable to load one or more of
% the requested types', it is usually caused by try to connect to a 32-bit
% version of OpticStudio from a 64-bit version of MATLAB (or vice-versa).
% This is an issue with how MATLAB interfaces with .NET, and the only
% current workaround is to use 32- or 64-bit versions of both applications.
app = TheConnection.ConnectAsExtension(instance);
if isempty(app)
   HandleError('Failed to connect to OpticStudio!');
end
if ~app.IsValidLicenseForAPI
	app.CloseApplication();
    HandleError('License check failed!');
    app = [];
end

end

function LogMessage(msg)
disp(msg);
end

function HandleError(error)
ME = MException('zosapi:HandleError', error);
throw(ME);
end



