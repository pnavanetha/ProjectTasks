import * as React from "react";
import { useState, useEffect } from "react";
import { DatePicker, DefaultButton, Dropdown, IDropdownOption, PrimaryButton, TextField } from '@fluentui/react';
import { getSP } from "../../pnpjsConfig";
// import { Item } from "@pnp/sp/items";

const TASK_LIST = 'ProjectTasksBacklogs';
const EMPLOYEE_LIST = 'ProjectEmployees';
const DOC_LIBRARY = 'ProjectTaskDocuments';
// const DOC_LOOKUP_ID = 'TaskRefId';

type Datekey =
    | 'EmailReceivedDate' | 'DevStartDate' | 'DevCompleteDate' | 'DevDeployedDate' | 'TestStartDate' | 'TestCompleteDate' | 'UATDeployedDate' | 'ProdDeploymentDate';

export interface ITaskForm extends Record<Datekey, Date | null> {
    Id?: number;
    Title: string;
    Priority: string;
    TaskType: string;
    EstimatedHours: string;
    ActualBurntHours: string;
    AssignedEmployeeId: number | null;
    AssignedEmployeeTitle?: string;
    Status: string;
    Remarks: string;
}

interface ITaskDocument {
    Id: number;
    FileLeafRef: string;
    FileRef: string;
}

const toOptions = (values: string[]): IDropdownOption[] =>
    values.map((v) => ({ key: v, text: v }));

const PRIORITY_OPTIONS = toOptions(['None', 'High', 'Medium', 'Low', 'Critical']);
const TYPE_OPTIONS = toOptions(['None', 'Analysis', 'Bug Fixing', 'Documentation', 'Deployment', 'Development', 'R&D', 'Testing', 'Test Cases', 'Unit Testing', 'User Manual', 'Coordination', 'Team Meeting', 'Requirement Clarification', 'Client Meetings', 'Not Required']);
const STATUS_OPTIONS = toOptions(['None', 'Ready', 'InProgress', 'Done', 'Blocked', 'On-Hold'])

const STAGES: { title: string; fields: { key: Datekey; label: string }[] }[] = [
    {
        title: 'Developmet',
        fields: [
            { key: 'DevStartDate', label: 'Start date' },
            { key: 'DevCompleteDate', label: 'Complete Date' },
            { key: 'DevDeployedDate', label: 'Deployed Date' }
        ]
    },
    {
        title: 'Testing',
        fields: [
            { key: 'TestStartDate', label: 'Start date' },
            { key: 'TestCompleteDate', label: 'Complete Date' }
        ]
    },
    {
        title: 'UAT',
        fields: [
            { key: 'UATDeployedDate', label: 'Deployed date' }
        ]
    },
    {
        title: 'Production',
        fields: [
            { key: 'ProdDeploymentDate', label: 'Deployment date' }
        ]
    },
];

const ALL_DATE_KEYS: Datekey[] = [
    'EmailReceivedDate',
    ...STAGES.reduce<Datekey[]>((acc, s) => acc.concat(s.fields.map((f) => f.key)), [])
];
const toSPDate = (d: Date | null): string | null =>
    d ? new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0).toISOString() : null;

const fromSPDate = (v: string | null | undefined): Date | null => (v ? new Date(v) : null);

const formatDate = (d: Date | null | undefined): string =>
    d ? d.toLocaleDateString('en-GB') : '';   // dd/mm/yyyy

const toNumber = (v: string): number | null => {
    const n = parseFloat(v);
    return isNaN(n) ? null : n;
};

const createEmptyForm = (): ITaskForm => {
    const dates = {} as Record<Datekey, Date | null>;
    ALL_DATE_KEYS.forEach((k) => { dates[k] = null; });
    return {
        ...dates,
        Title: '',
        Priority: '',
        TaskType: '',
        EstimatedHours: '',
        ActualBurntHours: '',
        AssignedEmployeeId: null,
        Status: '',
        Remarks: ''
    };
};
const cell: React.CSSProperties = { border: '1px solid #ccc', padding: '8px', textAlign: 'left' };
const sectionTitle: React.CSSProperties = { margin: '24px 0 8px', fontSize: 16, fontWeight: 600 };
const row: React.CSSProperties = { display: 'flex', flexWrap: 'wrap', gap: 16 };
const col: React.CSSProperties = { flex: '1 1 220px', minWidth: 200 };


const ProjectTaskForm: React.FC = () => {
    const [formData, setFormData] = useState<ITaskForm>(createEmptyForm());
    const [taskList, setTaskList] = useState<ITaskForm[]>([]);
    const [employees, setEmployees] = useState<IDropdownOption[]>([]);
    const [files, setFiles] = useState<File[]>([]);
    const [documents, setDocuments] = useState<ITaskDocument[]>([]);
    const [saving, setSaving] = useState<boolean>(false);

    const setField = <k extends keyof ITaskForm>(key: k, value: ITaskForm[k]): void => {
        setFormData((prev) => ({ ...prev, [key]: value }));
    };

    const getEmployees = async (): Promise<void> => {
        try {
            const sp = getSP();
            const items = await sp.web.lists.getByTitle(EMPLOYEE_LIST).items.select('Id', 'Title')();
            setEmployees(items.map((e) => ({ key: e.Id, text: e.Title })));
        } catch (error) {
            console.error('Error fetching employees:', error);
        }
    };

    const getTasks = async (): Promise<void> => {
        try {
            const sp = getSP();
            const items = await sp.web.lists.getByTitle(TASK_LIST).items.select('Id', 'Title', 'Priority', 'TaskType', 'EstimatedHours', 'ActualBurntHours', 'Status', 'Remarks', ...ALL_DATE_KEYS, 'AssignedEmployee/Id', 'AssignedEmployee/Title')
                .expand('AssignedEmployee').orderBy('Id', false)();
            const tasks: ITaskForm[] = items.map((item) => {
                const task = createEmptyForm();
                ALL_DATE_KEYS.forEach((k) => { task[k] = fromSPDate(item[k]); });
                return {
                    ...task,
                    Id: item.Id,
                    Title: item.Title || '',
                    Priority: item.Priority || '',
                    TaskType: item.TaskType || '',
                    EstimatedHours: item.EstimatedHours != null ? String(item.EstimatedHours) : '',
                    ActualBurntHours: item.ActualBurntHours != null ? String(item.ActualBurntHours) : '',
                    AssignedEmployeeId: item.AssignedEmployee ? item.AssignedEmployee.Id : null,
                    AssignedEmployeeTitle: item.AssignedEmployee ? item.AssignedEmployee.Title : '',
                    Status: item.Status || '',
                    Remarks: item.Remarks || ''
                };
            });
            setTaskList(tasks);

        } catch (error) {
            console.error('Error fetching tasks:', error);
        }
    };

    const getDocuments = async (taskId: number): Promise<void> => {
        try {
            const sp = getSP();
            const docs = await sp.web.lists
                .getByTitle(DOC_LIBRARY)
                .items
                .select('Id', 'FileLeafRef', 'FileRef', 'Title')
                .filter(`Title eq '${taskId}'`)();
            setDocuments(docs as ITaskDocument[]);
        } catch (error) {
            console.error('Error fetching documents:', error);
        }
    };

    useEffect(() => {
        getEmployees().catch(console.error);
        getTasks().catch(console.error);
    }, []);

    const buildPayload = (): Record<string, unknown> => {
        const payload: Record<string, unknown> = {
            Title: formData.Title,
            Priority: formData.Priority || null,
            TaskType: formData.TaskType || null,
            EstimatedHours: toNumber(formData.EstimatedHours),
            ActualBurntHours: toNumber(formData.ActualBurntHours),
            AssignedEmployeeId: formData.AssignedEmployeeId,   // lookup = <Column>Id
            Status: formData.Status || null,
            Remarks: formData.Remarks
        };
        ALL_DATE_KEYS.forEach((k) => { payload[k] = toSPDate(formData[k]); });
        return payload;
    };
    // ----- Upload files to the document library and link them to the task -----
    const uploadDocuments = async (taskId: number): Promise<void> => {
        if (files.length === 0) return;
        const sp = getSP();
        const library = sp.web.lists.getByTitle(DOC_LIBRARY);

        for (const file of files) {
            const uploaded = await library.rootFolder.files.addUsingPath(
                `${taskId}_${file.name}`, file, { Overwrite: true }
            );
            const item = await uploaded.file.getItem();
            await item.update({ Title: String(taskId) });
        }
    };

    const handleClear = (): void => {
        setFormData(createEmptyForm());
        setFiles([]);
        setDocuments([]);
    };

    const handleEdit = (task: ITaskForm): void => {
        setFormData({ ...task });
        setFiles([]);
        if (task.Id) {
            getDocuments(task.Id).catch(console.error);
        }
        window.scrollTo({ top: 0, behavior: 'smooth' });
    };

    const handleSave = async (): Promise<void> => {
        if (!formData.Title.trim()) {
            alert('Please enter the task name.');
            return;
        }
        setSaving(true);
        try {
            const sp = getSP();
            const result = await sp.web.lists
                .getByTitle(TASK_LIST)
                .items.add(buildPayload());

            const newId: number = result.data.Id;
            await uploadDocuments(newId);

            await getTasks();
            handleClear();
        } catch (error) {
            console.error('Error saving task:', error);
        } finally {
            setSaving(false);
        }
    };

    const handleUpdate = async (): Promise<void> => {
        if (!formData.Id) return;
        setSaving(true);
        try {
            const sp = getSP();
            await sp.web.lists
                .getByTitle(TASK_LIST)
                .items.getById(formData.Id)
                .update(buildPayload());

            await uploadDocuments(formData.Id);

            await getTasks();
            handleClear();
        } catch (error) {
            console.error('Error updating task:', error);
        } finally {
            setSaving(false);
        }
    };

    const renderDate = (key: Datekey, label: string): JSX.Element => (
        <div style={col} key={key}>
            <DatePicker
                label={label}
                placeholder="Select a date..."
                value={formData[key] || undefined}
                onSelectDate={(date) => setField(key, date || null)}
                formatDate={(date) => formatDate(date)}
                allowTextInput={false}
            />
        </div>
    );

    return (
        <div>
            <h3 style={sectionTitle}>Task details</h3>
            <div style={row}>
                <div style={{ ...col, flexBasis: '100%' }}>
                    <TextField label="Task" required value={formData.Title} onChange={(e, v) => setField('Title', v || '')} />
                </div>
                <div style={col}>
                    <Dropdown
                        label="Priority"
                        placeholder="Select priority"
                        options={PRIORITY_OPTIONS}
                        selectedKey={formData.Priority || null}
                        onChange={(_e, o) => setField('Priority', o ? String(o.key) : '')}
                    />
                </div>
                <div style={col}>
                    <Dropdown
                        label="Task type"
                        placeholder="Select type"
                        options={TYPE_OPTIONS}
                        selectedKey={formData.TaskType || null}
                        onChange={(_e, o) => setField('TaskType', o ? String(o.key) : '')}
                    />
                </div>
                <div style={col}>
                    <Dropdown
                        label="Assigned employee"
                        placeholder="Select employee"
                        options={employees}
                        selectedKey={formData.AssignedEmployeeId}
                        onChange={(_e, o) => setField('AssignedEmployeeId', o ? Number(o.key) : null)}
                    />
                </div>
                <div style={col}>
                    <Dropdown
                        label="Status"
                        placeholder="Select status"
                        options={STATUS_OPTIONS}
                        selectedKey={formData.Status || null}
                        onChange={(_e, o) => setField('Status', o ? String(o.key) : '')}
                    />
                </div>
                {renderDate('EmailReceivedDate', 'Email / task received date')}
                <div style={col}>
                    <TextField
                        label="Estimated hours"
                        type="number"
                        min={0}
                        step={0.5}
                        value={formData.EstimatedHours}
                        onChange={(_e, v) => setField('EstimatedHours', v || '')}
                    />
                </div>
                <div style={col}>
                    <TextField
                        label="Actual burnt hours"
                        type="number"
                        min={0}
                        step={0.5}
                        value={formData.ActualBurntHours}
                        onChange={(_e, v) => setField('ActualBurntHours', v || '')}
                    />
                </div>
            </div>

            {/* ---------- Stages ---------- */}
            {STAGES.map((stage) => (
                <div key={stage.title}>
                    <h3 style={sectionTitle}>{stage.title}</h3>
                    <div style={row}>
                        {stage.fields.map((f) => renderDate(f.key, f.label))}
                    </div>
                </div>
            ))}

            {/* ---------- Remarks ---------- */}
            <h3 style={sectionTitle}>Remarks</h3>
            <TextField
                multiline
                rows={3}
                value={formData.Remarks}
                onChange={(_e, v) => setField('Remarks', v || '')}
            />

            {/* ---------- Documents ---------- */}
            <h3 style={sectionTitle}>Documents</h3>
            <input
                type="file"
                multiple
                onChange={(e) => {
                    const selectedFiles = e.target.files
                        ? Array.prototype.slice.call(e.target.files) as File[]
                        : [];
                    setFiles(selectedFiles);
                }}
            />
            {files.length > 0 && (
                <div style={{ marginTop: 8 }}>
                    To upload: {files.map((f) => f.name).join(', ')}
                </div>
            )}
            {documents.length > 0 && (
                <ul style={{ marginTop: 8 }}>
                    {documents.map((doc) => (
                        <li key={doc.Id}>
                            <a href={doc.FileRef} target="_blank" rel="noreferrer">
                                {doc.FileLeafRef}
                            </a>
                        </li>
                    ))}
                </ul>
            )}

            {/* ---------- Buttons ---------- */}
            <div style={{ marginTop: 24, display: 'flex', gap: 8 }}>
                <PrimaryButton
                    text={saving ? 'Saving...' : formData.Id ? 'Update' : 'Save'}
                    disabled={saving}
                    onClick={() => {
                        (formData.Id ? handleUpdate() : handleSave()).catch(console.error);
                    }}
                />
                <DefaultButton text="Clear" disabled={saving} onClick={handleClear} />
            </div>

            {/* ---------- Table ---------- */}
            <div style={{ overflowX: 'auto', marginTop: 32 }}>
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                        <tr>
                            <th style={cell}>Edit</th>
                            <th style={cell}>ID</th>
                            <th style={cell}>Task Name</th>
                            <th style={cell}>Received Email Date</th>
                            <th style={cell}>Priority</th>
                            <th style={cell}>Task Type</th>
                            <th style={cell}>Assigned to</th>
                            <th style={cell}>Status</th>
                            <th style={cell}>Est. hrs</th>
                            <th style={cell}>Actual hrs</th>
                        </tr>
                    </thead>
                    <tbody>
                        {taskList.map((task) => (
                            <tr key={task.Id}>
                                <td style={cell}>
                                    <DefaultButton text="Edit" onClick={() => handleEdit(task)} />
                                </td>
                                <td style={cell}>{task.Id}</td>
                                <td style={cell}>{task.Title}</td>
                                <td style={cell}>{formatDate(task.EmailReceivedDate)}</td>
                                <td style={cell}>{task.Priority}</td>
                                <td style={cell}>{task.TaskType}</td>
                                <td style={cell}>{task.AssignedEmployeeTitle}</td>
                                <td style={cell}>{task.Status}</td>
                                <td style={cell}>{task.EstimatedHours}</td>
                                <td style={cell}>{task.ActualBurntHours}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>
    )


};
export default ProjectTaskForm;




