Attribute VB_Name = "QVb_Doc"
Option Explicit
Private Const CMod$ = "MVb_Doc."
Private Const Asm$ = "QVb"
Public Const DoczFunNmRul_1SubjFst$ = "If there is a Subj in pm, " & _
"put the Subj as fst CmlTerm and return that Subj; " & _
"give a Noun to the subj;" & _
"noun is MulCml."
Public Const DoczNoun$ = "It 1 or more Cml to form a Noun."
Public Const DoczTerm$ = "It a printable-string without space."
Public Const DoczSubj_1CmlTerm$ = "It is an a CmlTerm."
Public Const DoczSubj_2$ = "It is an instance of a Type."
Public Const DoczFunNmRul_1Type$ = "Each FunNm must belong to one of these rule: Noun | Noun.Verb.Extra | Verb.Variant | Noun.z.Variant"
Public Const DoczPmRul_1OneSubj$ = "Choose a subj in pm if there is more than one arg"
Public Const DoczPmRul_2MulNoun$ = "It is Ok to group mul-arg as one subj"
Public Const DoczPmRul_3MulNounUseOneCml$ = "Mul-noun as one subj use one Cml"
Public Const DoczCml$ = "Tag:Type. P1.NumIsLCase:.  P2.LowDashIsLCase:.  P3.FstChrCanAnyNmChr:."
Public Const DoczSfxSS = "Tag:NmRul. NmRul means variable or function name."
Public Const DoczSpecDocTerm_VdtVerbss$ = "Tag:Definition.  P1.Opt: Each module may one DoczVdtVerbss.  P2.OneOccurance: "
Public Const DoczNounVerbExtra$ = "Tag: FunNmRule.  Prp1.TakAndRetNoun: Fst Cml is Noun and Return Noun.  Prp2.OneCmlNoun: Noun should be 1 Cml.  " & _
"Prp3.VdtVerb: Snd Cml should be approved/valid noun.  Prp4.OptExtra: Extra is optional."


