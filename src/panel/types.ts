export type EditorEntityKind = 'threat' | 'securityObjective' | 'attacker' | 'model';

export interface ThreatEditorPayload {
  ID: string;
  title: string;
  attack: string;
  impactDesc: string;
  fullyMitigated: boolean;
  cvssVector: string;
  threatType: string;
  public: boolean;
}

export interface SecurityObjectiveEditorPayload {
  ID: string;
  title: string;
  description: string;
}

export interface AttackerEditorPayload {
  ID: string;
  name: string;
  title: string;
  description: string;
}

export interface ThreatModelEditorPayload {
  ID: string;
  title: string;
  version: string;
  analysis: string;
}

export type EntityEditorPayload = ThreatEditorPayload | SecurityObjectiveEditorPayload | AttackerEditorPayload | ThreatModelEditorPayload;

export interface EntityEditorViewModel {
  kind: EditorEntityKind;
  label: string;
  payload: EntityEditorPayload;
  focusField?: string;
}
