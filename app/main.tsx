import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  ExcelExport,
  ExcelExportColumn,
  KendoOoxml,
} from '@progress/kendo-react-excel-export';
import { Loader } from '@progress/kendo-react-indicators';
import { saveAs } from '@progress/kendo-file-saver';
import products from './productsLarge.json';

const data = products;
let cols = [
  'ProductID',
  'ProductName',
  'SupplierID',
  'CategoryID',
  'QuantityPerUnit',
  'UnitPrice',
  'UnitsInStock',
  'UnitsOnOrder',
  'ReorderLevel',
  'Discontinued',
  'ProductID',
  'ProductName',
  'SupplierID',
  'CategoryID',
  'QuantityPerUnit',
  'UnitPrice',
  'UnitsInStock',
  'UnitsOnOrder',
  'ReorderLevel',
  'Discontinued',
  'ProductID',
  'ProductName',
  'SupplierID',
  'CategoryID',
  'QuantityPerUnit',
  'UnitPrice',
  'UnitsInStock',
  'UnitsOnOrder',
  'ReorderLevel',
  'Discontinued',
  'ProductID',
  'ProductName',
  'SupplierID',
  'CategoryID',
  'QuantityPerUnit',
  'UnitPrice',
  'UnitsInStock',
  'UnitsOnOrder',
  'ReorderLevel',
  'Discontinued',
  'ProductID',
  'ProductName',
  'SupplierID',
  'CategoryID',
  'QuantityPerUnit',
  'UnitPrice',
  'UnitsInStock',
  'UnitsOnOrder',
  'ReorderLevel',
  'Discontinued',
  'ProductID',
  'ProductName',
  'SupplierID',
  'CategoryID',
  'QuantityPerUnit',
  'UnitPrice',
  'UnitsInStock',
  'UnitsOnOrder',
  'ReorderLevel',
  'Discontinued',
  'ProductID',
  'ProductName',
  'SupplierID',
  'CategoryID',
  'QuantityPerUnit',
  'UnitPrice',
  'UnitsInStock',
  'UnitsOnOrder',
  'ReorderLevel',
  'Discontinued',
  'ProductID',
  'ProductName',
  'SupplierID',
  'CategoryID',
  'QuantityPerUnit',
  'UnitPrice',
  'UnitsInStock',
  'UnitsOnOrder',
  'ReorderLevel',
  'Discontinued',
];
const App = () => {
  const [loading, setLoading] = React.useState(false);
  const _exporter = React.createRef();

  const save = async (component) => {
    const workbook = component.current.workbookOptions();
    const rows = workbook.sheets[0].rows;
    let altIdx = 0;
    rows.forEach((row) => {
      if (row.type === 'data' && altIdx++ % 2 !== 0) {
        row.cells.forEach((cell) => (cell.background = '#aabbcc'));
      }
    });
    setLoading(true);
    console.time('oxml');
    const dataUrl = await new KendoOoxml.Workbook(workbook).toDataURL();
    saveAs(dataUrl, 'Products.xlsx');
    console.timeEnd('oxml');
    setLoading(false);
  };

  const excelExport = () => save(_exporter);

  return (
    <div>
      <button
        className="k-button k-button-md k-rounded-md k-button-solid k-button-solid-base"
        onClick={excelExport}
      >
        Export to Excel
      </button>
      {loading && (
        <Loader
          size="large"
          style={{
            position: 'absolute',
            top: '50%',
            left: '50%',
            transform: 'translate(-50%,-50%)',
          }}
          type="infinite-spinner"
        />
      )}
      <ExcelExport data={data} fileName="Products.xlsx" ref={_exporter}>
        {cols.map((col, i) => (
          <ExcelExportColumn
            field={col}
            title={col.toUpperCase() + i}
            key={i}
          />
        ))}
      </ExcelExport>
    </div>
  );
};

ReactDOM.render(<App />, document.querySelector('my-app'));
