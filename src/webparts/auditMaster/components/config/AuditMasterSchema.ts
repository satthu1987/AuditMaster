
interface IFieldDefinition {
    key: string;
    aliases: string[];
    types: string[];
    includeInSelect?: boolean;
}
export const AUDIT_FIELD_CONFIG: IFieldDefinition[] = [
    {
        key: 'Service',
        aliases: ['service'],
        types: ['Text']
    },
    {
        key: 'Segment',
        aliases: ['segment'],
        types: ['Text']
    },
    {
        key: 'PIC',
        aliases: ['pic'],
        types: ['User', 'UserMulti']
    },
    {
        key: 'Auditor',
        aliases: ['auditor'],
        types: ['User', 'UserMulti']
    },
    {
        key: 'QualityManager',
        aliases: ['qualitymanager', 'qm', 'quality'],
        types: ['User', 'UserMulti']
    },
    {
        key: 'Verifier',
        aliases: ['verifier'],
        types: ['User', 'UserMulti']
    },
    {
        key: 'ISOChapter',
        aliases: [
            'isochapter'
        ],
        types: ['Lookup', 'LookupMulti'],
        includeInSelect: true
    },
    {
        key: 'RequiredRCA',
        aliases: [
            'requiredrca',
            'requiredrootcauseanalysis'
        ],
        types: ['Boolean'],
        includeInSelect: true
    },
    {
        key: 'QLVerification',
        aliases: [
            'qlverification',
            'qandlverification',
            'qualitylevelverification'
        ],
        types: [
            'Text'
        ],
        includeInSelect: true
    }
];
