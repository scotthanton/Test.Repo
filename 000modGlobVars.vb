'*******************************************************************************
'** Copyright by GasTOPS Ltd. 1979-2014. All rights reserved. No part of this
'** software may be reproduced in any form, or by any means, including photo-
'** copying, recording, taping, or information storage and retrieval systems,
'** except with written permission from GasTOPS Ltd.
'*******************************************************************************
'*******************************************************************************
'**                            FILE DESCRIPTION                               **
'*******************************************************************************
'**
'** File Name:    000modGlobVars.vb
'**
'** Class Name:   GlobalVariables
'**
'** Description:  Definition of common types used by this project.
'**
'** Methods:      
'**
'*******************************************************************************
'*******************************************************************************
'**                            REVISION HISTORY                               **
'*******************************************************************************
'**
'** $Log: 000modGlobVars.vb  $
'** Revision 1.8 2015/09/26 16:51:58GMT rtarb 
'** Replaced event logging using TraceSwitch with new mechanism using TraceSource.
'** Revision 1.7 2015/09/25 18:45:23GMT rtarb 
'** Added function-level instrumentation through the LogSWEvent function.
'** Revision 1.6 2015/09/24 13:57:18GMT rtarb 
'** Added access semaphore to protect function that writes to Windows Event Log.
'** Revision 1.5 2015/01/18 23:45:20GMT rtarb 
'** Changed definition of MS4K_Event class to replace Byte device address with
'** the unique device ID string.
'** Revision 1.4 2014/11/28 19:03:19GMT rtarb 
'** Added EventLogger class.
'** Revision 1.3 2014/11/19 20:33:00GMT rtarb 
'** Updated Message property of MS4K_Event class to change the way in which
'** messages are formatted if there is no associated value.
'** Revision 1.2 2014/11/17 15:56:03GMT rtarb 
'** Updated the Level property in the MS4K_Event class.
'**
'*******************************************************************************
'Define the options for code compilation/execution
Option Compare    Binary                                 'Text comparisons to be binary
Option Explicit   On                                     'All variables must be explicitly declared
Option Strict     Off

'Identify all namespaces imported by this module
Imports System.IO
Imports System.Threading

'Module definition
Public Module GlobalVariables

   '*******************************************************************************
   '** CONSTANTS USED BY THIS CLASS/FORM/MODULE
   '*******************************************************************************
   Public Const NULL_DATA  As Single   =  -9.99E+99      'Null data value
   Public Const NO_UNITS   As String   =  "[---]"        'No units of measurement

   '*******************************************************************************
   '** ENUMERATIONS USED BY THIS CLASS/FORM/MODULE
   '*******************************************************************************
   'Enumeration defining the event action to be taken
   Public Enum EventAction
      NONE        =  0           'No action
      LOGGED      =  1           'Event logged (new event)
      CLRD        =  2           'Event cleared
      ACKD        =  3           'Event acknowledged
   End Enum

   'Enumeration defining event severity for the MS4000 system
   Public Enum EventLevel
      SOFTWARE    = -1           'Software fault/error
      NONE        =  0           'No fault/warning/alarm
      INFO        =  0           'Information event
      FAULT       =  1           'Channel/device hardware fault
      WARNING     =  2           'Equipment health indicator is in warning
      ALARM       =  3           'Equipment health indicator is in alarm
   End Enum

   'Enumeration defining the MS4000 software operating modes
   Public Enum OperatingMode
      UNKNOWN     = -1
      AUTOMATIC   =  0
      MANUAL      =  1
   End Enum

   '*******************************************************************************
   '** GLOBAL VARIABLES
   '*******************************************************************************
   Public ev As EventLogger   =  New EventLogger( "GTL_MS3500.Utility" )

End Module 'End of GlobalVariables

'Definition of a class that facilitates the annunciation of an event
Public Class MS4K_Event 
   Inherits System.EventArgs

   '*******************************************************************************
   '** CLASS/FORM/MODULE ATTRIBUTES
   '*******************************************************************************
   '       Name            Type              Description                                     Log?
   '       ------------    --------          ----------------------------------------------  ----
   Private m_bAlert     As Boolean           'Event is to be displayed in the Alert Bar      Yes

   Private m_eAction    As EventAction       'Action to be taken with respect to logging     No
                                             'the event;

   Private m_eLevel     As EventLevel        'Event severity/level                           Yes

   Private m_sCID       As String            'Unique ID of channel/device that raised the    Yes
                                             'event; '0' for system info messages;

   Private m_sPointDesc As String            'Text description of DataPoint/measurement      No
                                             'that triggered the event; NULL for S/W faults 
                                             'and SYSINFO messages;

   Private m_sPointTag  As String            'Tag for DataPoint/measurement that triggered   Yes
                                             'the event; NULL for S/W faults and SYSINFO
                                             'messages;
   
   Private m_sMessage   As String            'Event message; S/W faults & SYSINFO messages   Yes
                                             'will set this value directly;

   Private m_sUnits     As String            'DataPoint units of measurement                 No

   Private m_sValue     As String            'DataPoint/measurement value that triggered     No
                                             'the event; NULL for S/W faults and SYSINFO
                                             'messages;

#Region "Public Control Properties"
   '*******************************************************************************
   '**
   '** PUBLIC CONTROL PROPERTIES
   '**
   '*******************************************************************************

   Public Property Action As EventAction
	'*******************************************************************************
	'**
	'** Property name:   Action
	'**
	'** Type:            EventAction
	'**
	'** Description:     (Property) Gets/sets the action code for the event.
	'**
	'*******************************************************************************
      Get
         'Retrieve the value of the class member attribute
         Action = m_eAction
      End Get
      Set ( value As EventAction )
         'Set the value of the class member attribute
         m_eAction = value
      End Set
   End Property 'End of Property Action

   Public Property Alert As Boolean
	'*******************************************************************************
	'**
	'** Property name:   Alert
	'**
	'** Type:            Boolean
	'**
	'** Description:     (Property) Gets/sets the value of the flag that determines
   '**                  whether logging the event alerts the GUI.
	'**
	'*******************************************************************************
      Get
         'Retrieve the value of the class member attribute
         Alert = m_bAlert
      End Get
      Set ( value As Boolean )
         'Set the value of the class member attribute
         m_bAlert = value
      End Set
   End Property 'End of Property Alert

   Public Property ChannelID As String
	'*******************************************************************************
	'**
	'** Property name:   ChannelID
	'**
	'** Type:            String
	'**
	'** Description:     (Property) Gets/sets the channel/device identifier of the
   '**                  hardware element associated with the event.
	'**
	'*******************************************************************************
      Get
         'Retrieve the value of the class member attribute
         ChannelID = m_sCID
      End Get
      Set ( value As String )
         'Set the value of the class member attribute
         m_sCID = value
      End Set
   End Property 'End of Property ChannelID

   Public Property Level As EventLevel
	'*******************************************************************************
	'**
	'** Property name:   Level
	'**
	'** Type:            EventLevel
	'**
	'** Description:     (Property) Gets/sets the severity code for the event.
	'**
	'*******************************************************************************
      Get
         'Return the current value of the class member attribute
         Level = m_eLevel
      End Get
      Set ( value As EventLevel )
         'Set the value of the class member attribute
         m_eLevel = value
      End Set
   End Property 'End of Property Level

   Public Property Message As String
	'*******************************************************************************
	'**
	'** Property name:   Message
	'**
	'** Type:            String
	'**
	'** Description:     (Property) Gets/sets the text message describing the event.
	'**
	'*******************************************************************************
      Get
         'If the message text has been specified...
         If Not String.IsNullOrEmpty( m_sMessage ) Then

            '...simply return the value of the class member attribute
            Message = m_sMessage
         Else

            'Initialize the function return value
            Message = ""

            '**********************************************************************
            '** FORMAT THE MESSAGE BASED ON THE EVENT INFORMATION
            '**********************************************************************
            'Start with the DataPoint Description...
            Message += m_sPointDesc + " "

            '...add the Event Level...
            Message += m_eLevel.ToString + " "

            '...add the Action Type...
            Message += m_eAction.ToString 

            'If the value text has been set...
            If Not String.IsNullOrEmpty( m_sValue ) Then

               '...add the DataPoint value to the message
               Message += IIf( Message.Length > 0, ": ", "" ) + "value = " + m_sValue

               '...and add the units of measurement, if required
               If Not String.IsNullOrEmpty( m_sUnits ) AndAlso m_sUnits <> NO_UNITS Then _
                  Message += " " + m_sUnits 

            End If 'End of If Not String.IsNullOrEmpty(...) block
         End If 'End of If Not String.IsNullOrEmpty(...) block
      End Get
      Set ( value As String )
         'Set the value of the class member attribute
         m_sMessage = value
      End Set
   End Property 'End of Property Message

   Public Property PointDesc As String
	'*******************************************************************************
	'**
	'** Property name:   PointDesc
	'**
	'** Type:            String
	'**
	'** Description:     (Property) Gets/sets the description of the DataPoint.
	'**
	'*******************************************************************************
      Get
         'Retrieve the value of the class member attribute
         PointDesc = m_sPointDesc
      End Get
      Set ( value As String )
         'Set the value of the class member attribute
         m_sPointDesc = value
      End Set
   End Property 'End of Property PointDesc

   Public Property PointTag As String
	'*******************************************************************************
	'**
	'** Property name:   PointTag
	'**
	'** Type:            String
	'**
	'** Description:     (Property) Gets/sets the tag of the DataPoint object that
   '**                  triggered the event.
	'**
	'*******************************************************************************
      Get
         'Retrieve the value of the class member attribute
         PointTag = m_sPointTag
      End Get
      Set ( value As String )
         'Set the value of the class member attribute
         m_sPointTag = value
      End Set
   End Property 'End of Property PointTag

   Public Property Units As String
	'*******************************************************************************
	'**
	'** Property name:   Units
	'**
	'** Type:            String
	'**
	'** Description:     (Property) Gets/sets the DataPoint units of measurement.
	'**
	'*******************************************************************************
      Get
         'Retrieve the value of the class member attribute
         Units = m_sUnits
      End Get
      Set ( value As String )
         'Set the value of the class member attribute
         m_sUnits = value
      End Set
   End Property 'End of Property Units

   Public Property Value As String
	'*******************************************************************************
	'**
	'** Property name:   Value
	'**
	'** Type:            String
	'**
	'** Description:     (Property) Gets/sets the DataPoint value.
	'**
	'*******************************************************************************
      Get
         'Retrieve the value of the class member attribute
         Value = m_sValue
      End Get
      Set ( value As String )
         'Set the value of the class member attribute
         m_sValue = value
      End Set
   End Property 'End of Property Value

#End Region

#Region "Public Class Functions"
   '*******************************************************************************
   '**
   '** PUBLIC CLASS FUNCTIONS
   '**
   '*******************************************************************************

    Public Sub New()
   '*******************************************************************************
   '**
   '** Routine Name: New
   '**
   '** Parameters:   None
   '**
   '** Description:  Default constructor for the MS4K_Event class.
   '**
   '** Returns:      Nothing
   '**
   '*******************************************************************************
      'Initialize the values of the class member attributes
      m_bAlert       =  False                         'Alert GUI when event logged?
      m_eAction      =  EventAction.LOGGED            'Event action
      m_eLevel       =  EventLevel.NONE               'Event level
      m_sMessage     =  ""                            'Event message
      m_sCID         =  "0"                           'Channel/device identifier
      m_sPointDesc   =  ""                            'DataPoint description
      m_sPointTag    =  ""                            'DataPoint identifier
      m_sUnits       =  NO_UNITS                      'Units of measurement
      m_sValue       =  ""                            'DataPoint value

   End Sub 'End of New

#End Region

End Class 'End of Class MS4K_Event

'End of file
 
