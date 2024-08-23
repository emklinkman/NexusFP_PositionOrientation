% Created by John Porter and Nev Pires
% Use this script to determine the position and orientation of one or more force plates
% relative to the lab coordinate system when direct measurement is impractical
% E.g. when the plates are located far from the volume origin, at a different level from the lab
% XY plane, inclined at unknown angles, etc.

% Force plate details, such as length and width and marker details, such as diameter, base thickness,
% base diameter, etc., are read from the ForcePlate.vst file which will need to be added to the session
% for each plate (subject) under measure. Default settings can be applied.

% 3 markers are to be placed on each force plate.
% MarkerA placed at force plate's local -X,-Y corner
% MarkerB placed at force plate's local +X,-Y corner
% MarkerC placed at force plate's local +X,+Y corner

% Force plate devices do not need to be added prior to running the script. The script will generate an
% XLSX file within the session folder which will contain a table listing the XYZ position and XYZ orientation
% for each force plate.

%%

clear
clc

%% Connect to Nexus
vicon = ViconNexus();

%% Get subject and marker information from trial
[ trialPath, trialName ] = vicon.GetTrialName();
[ subjectList, ~, active ] = vicon.GetSubjectInfo();

activeSubjectList = [ ];
for subject = 1 : length( subjectList ) % Build a subject list from those that are active in the session
    if active( subject )
        activeSubjectList = [ activeSubjectList, subjectList( subject ) ];
    end
end  

if isempty( activeSubjectList )
    
    uiwait( msgbox( 'No active subject found', 'modal' ) );
    return
    
end

activeSubjectList = sort( activeSubjectList );

for subject = 1 : length( activeSubjectList )
    
      markerList = vicon.GetMarkerNames( activeSubjectList{ subject } ); % Get the marker list
    
    if isempty( markerList )
    
        uiwait( msgbox( 'No marker set associated with the subject', 'modal' ) );
        return
    
        elseif length( markerList ) < 3
    
        uiwait( msgbox( 'The marker set requires at least three markers', 'modal' ) );
        return
    
    end
    
    hasData = TrajectoryCheck( vicon, ( activeSubjectList{ subject } ), markerList ); %Do all the trajectories even exist?
    
    if sum( hasData ) < length( markerList ) %Kick out if at least one marker is missing throughout the trial
        
        uiwait( msgbox( 'At least one marker is missing throughout the entire trial', 'modal' ) );
        return
             
    end
	
    %% Read force plate and marker dimensions from the force plate VSK file
    plateParams.( activeSubjectList{ subject } ).PlateDim_X = vicon.GetSubjectParam( ( activeSubjectList{ subject } ), 'PlateDim_X' );
    plateParams.( activeSubjectList{ subject } ).PlateDim_Y = vicon.GetSubjectParam( ( activeSubjectList{ subject } ), 'PlateDim_Y' );
    plateParams.( activeSubjectList{ subject } ).MarkerDiameter = vicon.GetSubjectParam( ( activeSubjectList{ subject } ), 'MarkerDiameter' );	
    plateParams.( activeSubjectList{ subject } ).MarkerBaseDiameter = vicon.GetSubjectParam( ( activeSubjectList{ subject } ), 'MarkerBaseDiameter' );
    plateParams.( activeSubjectList{ subject } ).MarkerBaseThickness = vicon.GetSubjectParam( ( activeSubjectList{ subject } ), 'MarkerBaseThickness' );		
    plateParams.( activeSubjectList{ subject } ).VerticalOffset = vicon.GetSubjectParam( ( activeSubjectList{ subject } ), 'VerticalOffset' );    
    
	%% Set reference marker local static coordinates (mm) in a structure
    refMarkers.( activeSubjectList{ subject } ).( 'MarkerA' ) = ...
        [ ( plateParams.( activeSubjectList{ subject } ).PlateDim_X - plateParams.( activeSubjectList{ subject } ).MarkerBaseDiameter ) / 2; ...
		( -plateParams.( activeSubjectList{ subject } ).PlateDim_Y + plateParams.( activeSubjectList{ subject } ).MarkerBaseDiameter ) / 2; ...
		plateParams.( activeSubjectList{ subject } ).MarkerBaseThickness + ( plateParams.( activeSubjectList{ subject } ).MarkerDiameter / 2 ) ...
		- plateParams.( activeSubjectList{ subject } ).VerticalOffset ];
    refMarkers.( activeSubjectList{ subject } ).( 'MarkerB' ) = ...
        [ ( -plateParams.( activeSubjectList{ subject } ).PlateDim_X + plateParams.( activeSubjectList{ subject } ).MarkerBaseDiameter ) / 2; ...
		( -plateParams.( activeSubjectList{ subject } ).PlateDim_Y + plateParams.( activeSubjectList{ subject } ).MarkerBaseDiameter ) / 2; ...
		plateParams.( activeSubjectList{ subject } ).MarkerBaseThickness + ( plateParams.( activeSubjectList{ subject } ).MarkerDiameter / 2 ) ...
		- plateParams.( activeSubjectList{ subject } ).VerticalOffset ];
    refMarkers.( activeSubjectList{ subject } ).( 'MarkerC' ) = ...
        [ ( -plateParams.( activeSubjectList{ subject } ).PlateDim_X + plateParams.( activeSubjectList{ subject } ).MarkerBaseDiameter ) / 2; ...
		( plateParams.( activeSubjectList{ subject } ).PlateDim_Y - plateParams.( activeSubjectList{ subject } ).MarkerBaseDiameter ) / 2; ...
		plateParams.( activeSubjectList{ subject } ).MarkerBaseThickness + ( plateParams.( activeSubjectList{ subject } ).MarkerDiameter / 2 ) ...
		- plateParams.( activeSubjectList{ subject } ).VerticalOffset ];

    %% Read trial marker data and place them in a structure and compile an average matrix
    FPMarkerData.( activeSubjectList{ subject } ) = GetMarkerData( vicon, activeSubjectList{ subject }, markerList );  
    FPMarkerData.( activeSubjectList{ subject } ) = SetNaN( FPMarkerData.( activeSubjectList{ subject } ) );
    avgFPMarkerMatrix.( activeSubjectList{ subject } ) = CreateAverageMatrix( FPMarkerData.( activeSubjectList{ subject } ) );
    FPRefMarkerMatrix.( activeSubjectList{ subject } ) = ...
        [refMarkers.( activeSubjectList{ subject } ).( 'MarkerA' ), refMarkers.( activeSubjectList{ subject } ).( 'MarkerB' ), refMarkers.( activeSubjectList{ subject } ).( 'MarkerC' ) ];
       
    % Use Singular Value Decomposition (SVD) to compute translation and rotation to move reference marker
    % cloud to global plate position marker cloud. Then compute global force plate position and orientaton
    % directly    
    
    [ Translate.( activeSubjectList{ subject } ), Rotate.( activeSubjectList{ subject } ) ] = ...
        ChangeMarkerCloudPose( FPRefMarkerMatrix.( activeSubjectList{ subject } ), avgFPMarkerMatrix.( activeSubjectList{ subject } ) );
    
    FP_Pos.( activeSubjectList{ subject } ) = Translate.( activeSubjectList{ subject } );
%        + Rotate.( subjectList{ subject } ) * [ ( plateParams.( subjectList{ subject } ).PlateDim_X ) / 2; ...
%        ( plateParams.( subjectList{ subject } ).PlateDim_Y ) / 2; 0 ];
    FP_Orient.( activeSubjectList{ subject } ) = AngleAxisFromMatrix( Rotate.( activeSubjectList{ subject } ) );    
    
end

%% Compile Force Plate Position and Orientation Output Data
    
PlatePositionOutput = zeros( 3, length( activeSubjectList ) );
PlateOrientationOutput = zeros( 3, length( activeSubjectList ) );

for subject = 1 : length( activeSubjectList )
    
    PlatePositionOutput( :, subject ) = FP_Pos.( activeSubjectList{ subject } );
    PlateOrientationOutput( :, subject ) = FP_Orient.( activeSubjectList{ subject } );
    
end  

%% Write Position and Orientation Output Data to Spreadsheet
trialName = replace( trialName, ' ', '_' );
outFileName = [ trialPath, trialName, '_Plate_Properties.xlsx' ];

xlswrite( outFileName, { 'Position' } , trialName, 'A1' )
xlswrite( outFileName, PlatePositionOutput, trialName, 'B2' )
xlswrite( outFileName, activeSubjectList, trialName, 'B1' )
xlswrite( outFileName, { 'x'; 'y'; 'z' }, trialName, 'A2' )

xlswrite( outFileName,{ 'Orientation' }, trialName, 'A5' )
xlswrite( outFileName, PlateOrientationOutput, trialName, 'B6' )
xlswrite( outFileName, activeSubjectList, trialName, 'B5' )
xlswrite( outFileName, { 'x'; 'y'; 'z' }, trialName, 'A6' )

% xlswrite( outFileName, { 'XLoc' } , trialName, 'B1' )
% xlswrite( outFileName, { 'YLoc' } , trialName, 'C1' )
% xlswrite( outFileName, { 'ZLoc' } , trialName, 'D1' )
% xlswrite( outFileName, PlatePositionOutput', trialName, 'B2' )
% xlswrite( outFileName, subjectList', trialName, 'A2' )
% 
% xlswrite( outFileName,{ 'XOrient' }, trialName, 'E1' )
% xlswrite( outFileName,{ 'YOrient' }, trialName, 'F1' )
% xlswrite( outFileName,{ 'ZOrient' }, trialName, 'G1' )
% xlswrite( outFileName, PlateOrientationOutput', trialName, 'E2' )


%% Functions

function DataExists = TrajectoryCheck( vicon, subject, markerlist )

DataExists = zeros( 1, length( markerlist ) );

for marker = 1 : length( markerlist )
    
    DataExists( :, marker ) = vicon.HasTrajectory( subject, char( markerlist( marker ) ) );
    
end

end

function MarkerArray = GetMarkerData( vicon, subject, markerlist )

numFrames = vicon.GetFrameCount;
[ SelectedStart, SelectedEnd ] = vicon.GetTrialRegionOfInterest();

% Read the marker data
MarkerArray = zeros( 4, length( markerlist ), numFrames );

    for marker = 1 : length( markerlist )
    
    [ x, y, z, e ] = vicon.GetTrajectory( subject, char( markerlist( marker ) ) );
    MarkerArray( :, marker, : ) = [ x; y; z; e ];
    
    end

    if numFrames > 1
    
       MarkerArray = MarkerArray( :, :, SelectedStart : SelectedEnd );
    
    end

end

function OutputArray = SetNaN( MarkerArray ) % Use only frames where all markers are present

    OutputArray = MarkerArray;
    
    for frame = 1 : size( MarkerArray, 3 )
        
      if sum( MarkerArray( 4, :, frame ) ) < size( MarkerArray, 2 )
          
         OutputArray( 1: 3, :, frame ) = NaN;
         OutputArray( 4, :, frame ) = false;
         
      end
      
    end
end

function MarkerMatrix = CreateAverageMatrix( MarkerArray ) % Construct a matrix from the average XYZ coordinates of each marker
    
    MarkerMatrix = zeros( 3, size( MarkerArray, 2 ) );
    
    for row = 1 : 3
        
        for column = 1 : size( MarkerArray, 2 )
            
             MarkerMatrix( row, column ) = mean( MarkerArray( row, column, : ), 'omitnan' );
             
        end
        
    end
    
end  

function [ T, R ] = ChangeMarkerCloudPose( InitialMatrix, FinalMatrix )

    Initial_Origin = mean( InitialMatrix, 2 ); % Centroid of markers
    Centered_Initial_Matrix = InitialMatrix - Initial_Origin; % Subtract the origin coordinates from each marker position
    Final_Origin = mean( FinalMatrix, 2 ); 
    Centered_Final_Matrix = FinalMatrix - Final_Origin;
    C = Centered_Final_Matrix * Centered_Initial_Matrix';
    
    [ U,~ ,V ] = svd( C ); % Use Singular Value Decomposition to decontruct C matrix into U and V matrix components
    
    detUVT = det( U * V' );
    L = eye( 3, 3 );
    
    if detUVT < 0
        
        L( 3, 3 ) = -1;
        
    end
    
    R = U * L * V'; % Compute rotation matrix that rotates marker cloud from initial orientation to final orientaton
    T = Final_Origin - R * Initial_Origin; % Compute translation vector that translates marker cloud from initial orientation to final orientation
    
end

function Orientation = AngleAxisFromMatrix( A ) % Compute the angle-axis form of the orientation from the rotation matrix

    angle = rad2deg( acos( ( A( 1, 1 ) + A( 2, 2 ) + A( 3, 3 ) - 1 ) / 2 ) );
    r = sqrt( ( A( 3, 2 ) - A( 2, 3 ) ) ^ 2+( A( 1, 3 ) - A( 3, 1 ) ) ^ 2+( A( 2, 1 ) - A( 1, 2 ) ) ^ 2);
    if r ~= 0
        x = ( A( 3, 2 ) - A( 2, 3 ) ) / r;
        y = ( A( 1, 3 ) - A( 3, 1 ) ) / r;
        z = ( A( 2, 1 ) - A( 1, 2 ) ) / r;
        Orientation = [ angle * x; angle * y; angle * z ];
    else
        Orientation = [ 0; 0; 0 ];
    end
    
end


%% 
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% automatically update .system file with force plate position & orientation
% added by EK Klinkman
% Version 2024.07.15

% Define the base path to the XML file
baseFilePath = 'your_file.system';

% Parse the XML content once
xml_content = fileread(baseFilePath);
xDoc = xmlreadstring(xml_content);

% Define the new values for the parameters for each force plate
new_values_array = {...
    containers.Map({'StandardPosition_X', 'StandardPosition_Y', 'StandardPosition_Z', ...
                    'StandardOrientation_X', 'StandardOrientation_Y', 'StandardOrientation_Z'}, ...
                   {'1.0', '2.0', '3.0', '4.0', '5.0', '6.0'}), ...
    containers.Map({'StandardPosition_X', 'StandardPosition_Y', 'StandardPosition_Z', ...
                    'StandardOrientation_X', 'StandardOrientation_Y', 'StandardOrientation_Z'}, ...
                   {'1.1', '2.1', '3.1', '4.1', '5.1', '6.1'}), ...
    containers.Map({'StandardPosition_X', 'StandardPosition_Y', 'StandardPosition_Z', ...
                    'StandardOrientation_X', 'StandardOrientation_Y', 'StandardOrientation_Z'}, ...
                   {'1.2', '2.2', '3.2', '4.2', '5.2', '6.2'}), ...
    containers.Map({'StandardPosition_X', 'StandardPosition_Y', 'StandardPosition_Z', ...
                    'StandardOrientation_X', 'StandardOrientation_Y', 'StandardOrientation_Z'}, ...
                   {'1.3', '2.3', '3.3', '4.3', '5.3', '6.3'}), ...
    containers.Map({'StandardPosition_X', 'StandardPosition_Y', 'StandardPosition_Z', ...
                    'StandardOrientation_X', 'StandardOrientation_Y', 'StandardOrientation_Z'}, ...
                   {'1.4', '2.4', '3.4', '4.4', '5.4', '6.4'})};

% Get all Param elements
allParams = xDoc.getElementsByTagName('Param');

% Iterate through the force plates
for i = 1:5
    % Get the current set of new values
    new_values = new_values_array{i};
    
    % Iterate over the parameters and update their values
    for k = 0:allParams.getLength-1
        param = allParams.item(k);
        nameAttr = param.getAttribute('name');
        
        if new_values.isKey(char(nameAttr))
            param.setAttribute('value', new_values(char(nameAttr)));
        end
    end
    
    % Save the modified XML to a new file for each iteration
    outputFilePath = sprintf('modified_file_force_plate_%d.system', i);
    xmlwrite(outputFilePath, xDoc);
end

disp('XML files updated successfully.');
