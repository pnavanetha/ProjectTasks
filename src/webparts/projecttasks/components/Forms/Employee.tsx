import * as React from 'react';
import { useEffect, useState } from 'react';
import { Checkbox, PrimaryButton, TextField } from '@fluentui/react';
import { getSP } from '../../pnpjsConfig';

const LIST_NAME = "ProjectEmployees";

export interface IEmpoyeeFormData {
    Id?: number;
    Title: string;
    isActive: boolean;
}

const EmployeeForm: React.FC = () => {
    const [formData, setFormData] = useState<IEmpoyeeFormData>({ Title: '', isActive: false });
    const [employeeList, setEmployeeList] = useState<IEmpoyeeFormData[]>([]);

    const getEmployees = async (): Promise<void> => {
        try {
            const sp = getSP();
            const items: IEmpoyeeFormData[] = await sp.web.lists
                .getByTitle(LIST_NAME)
                .items.select('Id', 'Title', 'isActive')
                ();
            setEmployeeList(items);

        } catch (error) {
            console.error("Error fetching employees:", error);

        }

    };

    useEffect(() => {
        getEmployees().catch(console.error);
    }, []);

    const handleEdit = (employee: IEmpoyeeFormData): void => {
        setFormData({ ...employee });
    };

    const handleClear = (): void => {
        setFormData({ Title: '', isActive: false });
    };

    const handleSave = async (): Promise<void> => {
        try {
            const sp = getSP();

            // SharePoint creates the Id automatically here
            const result = await sp.web.lists
                .getByTitle(LIST_NAME)
                .items
                .add({
                    Title: formData.Title,
                    IsActive: formData.isActive
                });
            console.log('Employee added:', result);
            // setEmployeeList([...employeeList, formData]);
            // setFormData({
            //     Title: '',
            //     isActive: false
            // });
            } catch (error) {
                console.error("Error saving employee:", error);
            }
        
    };

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
            <PrimaryButton text="Clear" styles={{ root: { backgroundColor: 'gray' } }} onClick={handleClear} />
            <table style={{ marginTop: '20px', width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                    <tr>
                        <th style={{ border: '1px solid black', padding: '8px' }}>Edit</th>
                        <th style={{ border: '1px solid black', padding: '8px' }}>ID</th>
                        <th style={{ border: '1px solid black', padding: '8px' }}>Employee Name</th>
                        <th style={{ border: '1px solid black', padding: '8px' }}>Is Active</th>
                    </tr>
                </thead>
                <tbody>
                    {employeeList.map((employee, index) => (
                        <tr key={index}>
                            <td style={{ border: '1px solid black', padding: '8px' }}>
                                <PrimaryButton text="Edit" onClick={() => handleEdit(employee)} />
                            </td>
                            <td style={{ border: '1px solid black', padding: '8px' }}>{employee.Id}</td>
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

