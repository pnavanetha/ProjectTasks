import * as React from 'react';
import { useState } from 'react';
import { Checkbox, PrimaryButton, TextField } from '@fluentui/react';

export interface IEmpoyeeFormData {
    id?: number;
    Title: string;
    isActive: boolean;
}

const EmployeeForm: React.FC = () => {
    const [formData, setFormData] = useState<IEmpoyeeFormData>({
        Title: '',
        isActive: false
    });
    const [employeeList, setEmployeeList] = useState<IEmpoyeeFormData[]>([]);

    const handleSave = () => {
        console.log('Form Data:', formData);
        setEmployeeList([...employeeList, formData]);
 
        setFormData({
            Title: '',
            isActive: false
        });
    }

    return (
        <div>
            <TextField
                label="Employee Name"
                value={formData.Title}
                onChange={(_event, newValue) => setFormData({ ...formData, Title: newValue || '' })}
            /> <br />
            <Checkbox
                label="Is Active"
                checked={formData.isActive}
                onChange={(_event, checked) => setFormData({ ...formData, isActive: checked || false })}
            /> <br />
            <PrimaryButton text="Save" onClick={handleSave} />
            <table style={{ marginTop: '20px', width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                    <tr>
                        <th style={{ border: '1px solid black', padding: '8px' }}>Title</th>
                        <th style={{ border: '1px solid black', padding: '8px' }}>Is Active</th>
                    </tr>
                </thead>
                <tbody>
                    {employeeList.map((employee, index) => (
                        <tr key={index}>
                            <td style={{ border: '1px solid black', padding: '8px' }}>{employee.Title}</td>
                            <td style={{ border: '1px solid black', padding: '8px' }}>{employee.isActive ? 'Yes' : 'No'}</td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};

export default EmployeeForm;

