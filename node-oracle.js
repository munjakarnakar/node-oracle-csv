const oracledb = require('oracledb');
const Excel = require('exceljs');
const fs = require("fs");

const db_user_name = "demonode"
const db_password = "demonode"
const db_connection_url = "demonode"

async function runApp() {
    let connection;
    try {
        connection = await oracledb.getConnection({
            user: db_user_name,
            password: db_password,
            connectionString: db_connection_url
        });

        console.log("Successfully connected to Oracle Database");

        // Now query the rows back
        // result = await connection.execute(`select description, done from todoitem`, [], { resultSet: true, outFormat: oracledb.OUT_FORMAT_OBJECT });
        const result = { resultSet: [{ name: 'karnakar', phone: '9000423012' }] }
        const rs = result.resultSet;
        console.log('Result set', rs);
        const final_result = [];
        let row;
        while ((row = await rs.getRow())) {
            final_result.push(final_result)
            console.log(row);
        }

        const headers = [];
        const keys = Object.keys(final_result[0]);

        keys.map((key) => headers.push({ header: key.replace(/_/g, ' ').toUpperCase(), key }));
        await generateExcel(headers, final_result, `CSV_${final_result.length}`);
        await rs.close();

    } catch (err) {
        console.error(err);
    } finally {
        if (connection) {
            try {
                await connection.close();
            } catch (err) {
                console.error(err);
            }
        }
    }
}

async function generateExcel(headers, data) {
    const file_name = `${data.length}-streamed-workbook.xlsx`;
    const reports_path = `${__dirname}/CSV_FILES`;
    if (!fs.existsSync(reports_path)) {
        // If it doesn't exist, create it
        fs.mkdirSync(reports_path, { recursive: true });
        console.log({ descrisendingption: 'Excel-Test:Folder created successfully', jsonObject: { reports_path } });
    } else {
        console.log({ description: 'Excel-Test:Folder already exists', jsonObject: { reports_path } });
    }

    const file_path = `${reports_path}/${file_name}`;

    const options = {
        filename: file_path,
        useStyles: true,
        useSharedStrings: true
    };

    const workbook = new Excel.stream.xlsx.WorkbookWriter(options);

    console.log({ description: 'Excel:Workbook created', jsonObject: {} });

    const worksheet = workbook.addWorksheet('Data');

    worksheet.columns = headers;

    console.log({ description: 'Excel:Added columns to work book', jsonObject: { headers } });

    // console.log(`Writing data to workbook: **Columns ${headers.length}**`, `**Rows ${data.length}**`);

    for (let index = 0; index < data.length; index++) {
        console.log(index, data[index])
        worksheet.addRow(data[index]).commit();
    }

    await workbook.commit();

    console.log({ description: 'Excel:Excel generated in temp path:', jsonObject: file_path });

    return { file_path, file_name };
};
runApp();