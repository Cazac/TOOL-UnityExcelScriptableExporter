using System.Collections;
using System.Collections.Generic;
using UnityEngine;

[System.Serializable]
[CreateAssetMenu(fileName = "Event", menuName = "Scriptables/New Event")]
public class RARC_Event_SO : ScriptableObject
{
    ////////////////////////////////

    [Header("Event Info")]
    public string eventTitle;
    public string eventID;
    public Sprite eventIcon;

    [TextArea()]
    public string eventDescription;

    [Header("Option 1 (Null = No Choice)")]
    [TextArea()]
    public string eventOption1_Choice;
    public RARC_EventOutcome_SO eventOption1_Outcome;
    public RARC_EventRequirement_SO eventOption1_Requirement;

    [Header("Option 2 (Null = No Choice)")]
    [TextArea()]
    public string eventOption2_Choice;
    public RARC_EventOutcome_SO eventOption2_Outcome;
    public RARC_EventRequirement_SO eventOption2_Requirement;

    [Header("Option 3 (Null = No Choice)")]
    [TextArea()]
    public string eventOption3_Choice;
    public RARC_EventOutcome_SO eventOption3_Outcome;
    public RARC_EventRequirement_SO eventOption3_Requirement;

    [Header("Allow the player to come back")]
    public bool eventCanComeBackLater;

    /////////////////////////////////////////////////////////////////
}
