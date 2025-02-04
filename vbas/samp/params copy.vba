' Create Parameters for switching on-off selected solids

'--------------------------------------------------------------------------------------------------------------------------------------------------------
' 18-Jun-2013 ube: macro is now also working for nested component folders, e.g. component1/component2/solid1
' 12-Jun-2013 tgl/ube: the old macro only worked for hex mesh, the new macro also works for tet, it is setting solid to vacuum instead of hiding it from simulation
' 11-Mar-2011 fsr: updated to vba_globals_all.lib+vba_globals_3D.lib, fixed an issue with opening the output file,
'					included Component Name so that solids With identical names under different components can be handled,
'					column width of parameter file is now determined dynamically
' 21-Jul-2009 ube: MakeSureParameterExists now written in history, fast model update and dependencies are properly recognized
' 20-Mar-2009 ube: always switch off Solid.FastModelUpdate
' 19-Mar-2009 ube: first version
'--------------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

'#include "vba_globals_all.lib"
'#include "vba_globals_3D.lib"

Sub Main ()
	With ParameterSweep
        .AddSequence "Sequence 1"
        .AddSequence "Sequence 2"
        .AddSequence "Sequence 3"
        .AddSequence "Sequence 4"
        .AddSequence "Sequence 5"
        .AddSequence "Sequence 6"
        .AddSequence "Sequence 7"
        .AddSequence "Sequence 8"
        .AddSequence "Sequence 9"
    End With
End Sub