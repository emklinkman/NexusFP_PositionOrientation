# Vicon Nexus Force Plate Position and Orientation
EK Klinkman

Neurobionics & Locomotor Control Systems Laboratories, Department of Robotics, University of Michigan, 2024

(MATLAB script written by Nev Pires & John Porter, Vicon)

## Overview

This project uses two scripts intended to process and manage force plate position and orientation in Vicon Nexus using MATLAB and Python. This process automates an otherwise manual process of updating position/orientation. It was written for the NeuroLoco labs at the University of Michigan Robotics department, where we have an adjustable circuit with instrumented stairs. If the circuit platform height is changed, the stair position and orientation will no longer be accurate, thus we developed a system to update it automatically .

## Files

1. 'Force_Plate_Position_and_Orientation_Generic_v2.m'
   * Provided by Vicon support engineers (Nev Pires, John Porter)
   * Pulls position and orientation of any specified force plate within the Vicon Nexus environment and writes these values to an Excel file.
   * NOT included in this repository
   
2. 'Vicon_XML_write_v5.py'
   * Written by NeuroLoco team (EK Klinkman)
   * Reads Excel file output data and writes a new .system XML file to be read into Nexus to automatically update force plate position and orientation.
  
3. 'Auto Label Stairs (EKK).System'
   * Vicon Nexus .System configuration base file example

4. 'Auto Label Stairs (EKK)_appended.System'
   * Vicon Nexus .System configuration appended file example
   * This is the result of running 'Force_Plate_Position_and_orientation_Generic_v2.m' and "Vicon_XML_write_v5.py'
   
## Getting Started

### Prerequisites
    * Python: Ensure Python is installed on your machine to run 'Vicon_XML_write.py'.
	* MATLAB: MATLAB is required to execute 'Force_Plate_Position_and_Orientation_Generic_v2.m'.
	
