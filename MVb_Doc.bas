Attribute VB_Name = "MVb_Doc"
Option Explicit
Public Const DocOfFunNmRul_1SubjFst$ = "If there is a Subj in pm, " & _
"put the Subj as fst CmlTerm and return that Subj; " & _
"give a Noun to the subj;" & _
"noun is MulCml."
Public Const DocOfNoun$ = "It 1 or more Cml to form a Noun."
Public Const DocOfTerm$ = "It a printable-string without space."
Public Const DocOfSubj_1CmlTerm$ = "It is an a CmlTerm."
Public Const DocOfSubj_2$ = "It is an instance of a Type."
Public Const DocOfFunNmRul_1Type$ = "Each FunNm must belong to one of these rule: Noun | Noun.Verb.Extra | Verb.Variant | Noun.z.Variant"
Public Const DocOfPmRul_1OneSubj$ = "Choose a subj in pm if there is more than one arg"
Public Const DocOfPmRul_2MulNoun$ = "It is Ok to group mul-arg as one subj"
Public Const DocOfPmRul_3MulNounUseOneCml$ = "Mul-noun as one subj use one Cml"
Public Const DocOfCml$ = "Tag:Type. P1.NumIsLCase:.  P2.LowDashIsLCase:.  P3.FstChrCanAnyNmChr:."
Public Const DocOfSfxSS = "Tag:NmRul. NmRul means variable or function name."
Public Const DocOfSpecDocTerm_VdtVerbss$ = "Tag:Definition.  P1.Opt: Each module may one DocOfVdtVerbss.  P2.OneOccurance: "
Public Const DocOfNounVerbExtra$ = "Tag: FunNmRule.  Prp1.TakAndRetNoun: Fst Cml is Noun and Return Noun.  Prp2.OneCmlNoun: Noun should be 1 Cml.  " & _
"Prp3.VdtVerb: Snd Cml should be approved/valid noun.  Prp4.OptExtra: Extra is optional."


