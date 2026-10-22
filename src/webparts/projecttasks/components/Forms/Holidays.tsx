import * as React from 'react';
import { useEffect, useState } from 'react';
import { DatePicker, PrimaryButton, TextField } from '@fluentui/react';
import { getSP } from '../../pnpjsConfig';

const LIST_NAME = 'EmployeeHolidays';

export interface IholidayHolidaysFormData {
    Id?: number;
    Title: string;
    EventDate: Date | null;
}


const toSPDate = (d: Date | null): string | null =>
    d ? new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0).toISOString() : null;

const formatDate = (d: Date | null): string =>
    d ? d.toLocaleDateString('en-GB') : '';   // dd/mm/yyyy

const HolidayForm: React.FC = () => {
    const [formData, setFormData] = useState<IholidayHolidaysFormData>({ Title: '', EventDate: null });
    const [holidayList, setHolidayList] = useState<IholidayHolidaysFormData[]>([]);

    const getHolidays = async (): Promise<void> => {
        try {
            const sp = getSP();
            const items = await sp.web.lists
                .getByTitle(LIST_NAME)
                .items.select('Id', 'Title', 'EventDate')();

            const holidays: IholidayHolidaysFormData[] = items.map((item) => ({
                Id: item.Id,
                Title: item.Title,
                EventDate: item.EventDate ? new Date(item.EventDate) : null
            }));
            setHolidayList(holidays);
        } catch (error) {
            console.error('Error fetching holidays:', error);
        }
    };

    useEffect(() => {
        getHolidays().catch(console.error);
    }, []);

    const handleEdit = (holiday: IholidayHolidaysFormData): void => {
        setFormData({ ...holiday });
    };

    const handleClear = (): void => {
        setFormData({ Title: '', EventDate: null });
    };

    const handleSave = async (): Promise<void> => {
        try {
            const sp = getSP();
            const result = await sp.web.lists
                .getByTitle(LIST_NAME)
                .items
                .add({
                    Title: formData.Title,
                    EventDate: toSPDate(formData.EventDate)
                });
            console.log('Holiday added:', result);

            await getHolidays();   // reload the table
            handleClear();         // clear the form
        } catch (error) {
            console.error('Error saving holiday:', error);
        }
    };

    const handleUpdate = async (): Promise<void> => {
        if (!formData.Id) return;
        try {
            const sp = getSP();
            await sp.web.lists
                .getByTitle(LIST_NAME)
                .items
                .getById(formData.Id)
                .update({
                    Title: formData.Title,
                    EventDate: toSPDate(formData.EventDate)
                });

            await getHolidays();
            handleClear();
        } catch (error) {
            console.error('Error updating holiday:', error);
        }
    };

    return (
        <div>
            <TextField
                label="Holiday Name"
                value={formData.Title}
                onChange={(_event, newValue) => setFormData({ ...formData, Title: newValue || '' })}
            /> <br />
            <DatePicker
                label="Date"
                placeholder="Select a date..."
                value={formData.EventDate || undefined}
                onSelectDate={(date) => setFormData({ ...formData, EventDate: date || null })}
                formatDate={(date) => formatDate(date || null)}
                allowTextInput={false}
            /> <br />

            <PrimaryButton
                text={formData.Id ? 'Update' : 'Save'}
                onClick={() => {
                    (formData.Id ? handleUpdate() : handleSave()).catch(console.error);
                }}
            />
            <PrimaryButton
                text="Clear"
                styles={{ root: { backgroundColor: 'gray', marginLeft: 8 } }}
                onClick={handleClear}
            />

            <table style={{ marginTop: '20px', width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                    <tr>
                        <th style={{ border: '1px solid black', padding: '8px' }}>Edit</th>
                        <th style={{ border: '1px solid black', padding: '8px' }}>Holiday Name</th>
                        <th style={{ border: '1px solid black', padding: '8px' }}>Event Date</th>
                    </tr>
                </thead>
                <tbody>
                    {holidayList.map((holiday) => (
                        <tr key={holiday.Id}>
                            <td style={{ border: '1px solid black', padding: '8px' }}>
                                <PrimaryButton text="Edit" onClick={() => handleEdit(holiday)} />
                            </td>
                            <td style={{ border: '1px solid black', padding: '8px' }}>{holiday.Title}</td>
                            <td style={{ border: '1px solid black', padding: '8px' }}>{formatDate(holiday.EventDate)}</td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};

export default HolidayForm;