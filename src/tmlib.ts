
import fs from 'fs';
import YAML, { Document } from 'yaml';


interface CVSS {
    vector: string
}
interface CounterMeasure {
    ID: string
    title: string
    description: string
    inPlace: boolean
    public: boolean
}    

interface Threat {
    ID: string
    title: string
    //attackers:
    //  - REFID: ANONYMOUS
    attack: string
    impactDesc: string
    
    fullyMitigated: boolean
    CVSS: CVSS
    threatType: string
    public: boolean
    
    countermeasures: CounterMeasure[]
}
interface SecurityObjective {
    ID?: string
    title?: string
    description?: string
    [key: string]: unknown
}

interface Attacker {
    ID?: string
    title?: string
    name?: string
    [key: string]: unknown
}

    interface Scope {
    description: string
}

export interface ThreatModel {
    ID: string
    title?: string
    parent?: string
    children?: { REFID: string }[]
    scope: Scope
    analysis: string
    threats: Threat[]
    securityObjectives?: SecurityObjective[]
    attackers?: Attacker[]
}

export function loadThreatModel(path: string) : ThreatModel{
    const file = fs.readFileSync(path, 'utf8');
    return parseThreatModelFromString(file);
}

export function parseThreatModelFromString(source: string): ThreatModel {
    const y = YAML.parseDocument(source, {keepSourceTokens:true});
    return y.toJS() as ThreatModel;
}

export function loadYAML(path: string) : Document {
    const file = fs.readFileSync(path, 'utf8');
    return parseYAMLFromString(file);
}

export function parseYAMLFromString(source: string): Document {
    const y = YAML.parseDocument(source, {keepSourceTokens:true});
    return y;
}
export function threats(model: any) {

}
