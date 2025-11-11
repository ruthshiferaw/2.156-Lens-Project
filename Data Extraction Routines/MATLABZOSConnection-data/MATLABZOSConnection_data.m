function [ TheApplication ] = MATLABZOSConnection_data( instance )

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
    

    % %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % % Custom Code: Export Lens Data to CSV (2024+ compatible)
    % %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    % %%% MY ATTEMPT %%%
    % TheSystem = TheApplication.PrimarySystem;
    % LDE = TheSystem.LDE;
    % 
    % % methods(LDE)
    % % properties(LDE)
    % 
    % surf = LDE.GetSurfaceAt(11);
    % % methods(surf)
    % % properties(surf);
    % 
    % settings = surf.CurrentTypeSettings();
    % % asp = surf.Type();
    % data = surf.SurfaceData()
    % % methods(data)
    % order4 = data.GetNthEvenOrderTerm(4)

    TheSystem = TheApplication.PrimarySystem;
    LDE = TheSystem.LDE;
    numSurf = LDE.NumberOfSurfaces;
    
    % Desired even orders to probe (A2, A4, ... A16)
    even_orders = 2:2:16;
    nEven = numel(even_orders);
    even_colnames = arrayfun(@(o) sprintf('A%d', o), even_orders, 'UniformOutput', false);
    
    % Prepare storage (use cell for mixed types)
    rows = cell(numSurf, 8 + nEven); % columns: Surface, TypeName, Comment, Radius, Thickness, Material, SemiDiameter, Conic, A2..A16
    headers = [{'Surface','TypeName','Comment','Radius','Thickness','Material','SemiDiameter','Conic'}, even_colnames];
    
    for idx = 1:numSurf
        try
            surf = LDE.GetSurfaceAt(idx-1); % zero-based in ZOS-API
        catch ME
            warning('Could not GetSurfaceAt(%d): %s', idx-1, ME.message);
            continue;
        end
    
        % default values
        s_surface = idx-1;
        s_type = '';
        s_comment = '';
        s_radius = NaN;
        s_thickness = NaN;
        s_material = '';
        s_semi = NaN;
        s_conic = NaN;
        s_even = nan(1,nEven);
    
        % Basic properties (many of these may throw on special rows, so wrap)
        try, s_type = char(surf.TypeName); end
        try, s_comment = char(surf.Comment); end
        try, s_radius = double(surf.Radius); end
        try, s_thickness = double(surf.Thickness); end
        try, s_material = char(surf.Material); end
        try, s_semi = double(surf.SemiDiameter); end
        try, s_conic = double(surf.Conic); end
    
        % Try to access SurfaceData (different concrete classes for types)
        surfData = [];
        try
            surfData = surf.SurfaceData;  % property, not a call
        catch
            surfData = [];
        end
    
        if ~isempty(surfData)
            % Get type name of the SurfaceData object
            try
                sdType = char(surfData.GetType().Name);
            catch
                sdType = '';
            end
    
            % 1) If object advertises GetNthEvenOrderTerm, use it (some ZOS versions)
            try
                m = methods(surfData);
                if any(strcmp(m, 'GetNthEvenOrderTerm'))
                    % We'll call GetNthEvenOrderTerm for each even order.
                    % Some versions accept the actual order (e.g., 4 -> 4th order)
                    for k = 1:nEven
                        order = even_orders(k);
                        try
                            val = surfData.GetNthEvenOrderTerm(order); % may return numeric
                            s_even(k) = double(val);
                        catch
                            % fallback: perhaps the method expects index (1->A2,2->A4,...)
                            try
                                val = surfData.GetNthEvenOrderTerm(k);
                                s_even(k) = double(val);
                            catch
                                s_even(k) = NaN;
                            end
                        end
                    end
                % 2) If object exposes Par1..ParN fields (common pattern)
                elseif any(startsWith(m, 'get_Par')) || any(contains(m, 'Par'))
                    % try reading Par1..ParN via property access
                    % Determine number of Par fields by probing
                    for k = 1:nEven
                        pname = sprintf('Par%d', k);
                        try
                            pcell = surfData.(pname); % ZOSAPI_LDECell-like
                            % cell might have DoubleValue or Value or StringValue
                            try
                                s_even(k) = double(pcell.DoubleValue);
                            catch
                                try
                                    s_even(k) = double(pcell.Value);
                                catch
                                    try
                                        s_even(k) = str2double(char(pcell.StringValue));
                                    catch
                                        s_even(k) = NaN;
                                    end
                                end
                            end
                        catch
                            s_even(k) = NaN;
                        end
                    end
                % 3) If object has NumberOfOrders and GetCoefficient or similar
                elseif any(strcmp(m, 'GetCoefficient')) || any(strcmp(m, 'GetNthTerm')) || any(strcmp(m, 'GetTerm'))
                    % Attempt to call a generic getter
                    for k = 1:nEven
                        try
                            % Some APIs use coefficient index starting at 1
                            val = surfData.GetCoefficient(k); %#ok<NASGU>
                            s_even(k) = double(val);
                        catch
                            try
                                val = surfData.GetCoefficient(even_orders(k)); % try using order
                                s_even(k) = double(val);
                            catch
                                s_even(k) = NaN;
                            end
                        end
                    end
                else
                    % No recognized API - leave NaNs
                end
            catch
                % any unexpected failure -> leave s_even as NaNs
            end
        end
    
        % Build row
        rowcell = cell(1, 8 + nEven);
        rowcell{1} = s_surface;
        rowcell{2} = s_type;
        rowcell{3} = s_comment;
        rowcell{4} = s_radius;
        rowcell{5} = s_thickness;
        rowcell{6} = s_material;
        rowcell{7} = s_semi;
        rowcell{8} = s_conic;
        for k = 1:nEven
            rowcell{8 + k} = s_even(k);
        end
    
        rows(idx, :) = rowcell;
    end
    
    % Convert cell array to table with headers
    T = cell2table(rows, 'VariableNames', headers);
    
    % Replace empty strings in numeric columns with NaN (if needed)
    for c = 4:8+nEven
        % if the column is not numeric, skip conversion
        if ~isnumeric(T{1,c})
            % attempt to convert cells that may contain numeric strings
            try
                T{:,c} = cellfun(@(x) (isempty(x) && isnumeric(x)) * NaN + (ischar(x) && ~isempty(x) * str2double(x)) , T{:,c});
            catch
                % ignore if conversion fails
            end
        end
    end
    
    % Write out CSV
    outPath = fullfile(getenv('USERPROFILE'), 'Documents', 'lens_data_all_surfaces.csv');
    writetable(T, outPath);
    fprintf('Exported %d surfaces to %s\n', height(T), outPath);

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
