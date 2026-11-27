import { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { v4 as uuidv4 } from 'uuid';
import {
  FileSpreadsheet,
  Upload,
  Settings,
  Play,
  Download,
  AlertTriangle,
  CheckCircle,
  Info,
  X,
  Trash2,
} from 'lucide-react';

// Types
interface SheetData {
  headers: string[];
  rows: Record<string, unknown>[];
}

interface LoadedFile {
  id: string;
  name: string;
  data: SheetData;
}

interface LogEntry {
  type: 'info' | 'success' | 'warning' | 'error';
  message: string;
  timestamp: Date;
}

// Main App Component
export default function App() {
  const [template, setTemplate] = useState<LoadedFile | null>(null);
  const [reports, setReports] = useState<LoadedFile[]>([]);
  const [keyColumns, setKeyColumns] = useState<string[]>([]);
  const [targetColumns, setTargetColumns] = useState<string[]>([]);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [resultData, setResultData] = useState<SheetData | null>(null);
  const [step, setStep] = useState<1 | 2 | 3 | 4>(1);
  const [isProcessing, setIsProcessing] = useState(false);

  const addLog = useCallback((type: LogEntry['type'], message: string) => {
    setLogs((prev) => [...prev, { type, message, timestamp: new Date() }]);
  }, []);

  const clearLogs = () => setLogs([]);

  // Parse Excel file
  const parseExcelFile = (file: File): Promise<SheetData> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, {
            defval: '',
          });

          if (jsonData.length === 0) {
            reject(new Error('Файл пуст или не содержит данных'));
            return;
          }

          const headers = Object.keys(jsonData[0]);
          resolve({ headers, rows: jsonData });
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  // Load template file
  const handleTemplateUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const data = await parseExcelFile(file);
      setTemplate({ id: uuidv4(), name: file.name, data });
      setKeyColumns([]);
      setTargetColumns([]);
      setResultData(null);
      setLogs([]);
      addLog('success', `Шаблон "${file.name}" загружен. Столбцов: ${data.headers.length}, строк: ${data.rows.length}`);
      if (reports.length > 0) {
        setStep(2);
      } else {
        setStep(2);
      }
    } catch (err) {
      addLog('error', `Ошибка загрузки шаблона: ${(err as Error).message}`);
    }
    e.target.value = '';
  };

  // Load report files
  const handleReportsUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    const newReports: LoadedFile[] = [];
    for (const file of files) {
      try {
        const data = await parseExcelFile(file);
        newReports.push({ id: uuidv4(), name: file.name, data });
        addLog('success', `Отчёт "${file.name}" загружен. Столбцов: ${data.headers.length}, строк: ${data.rows.length}`);
      } catch (err) {
        addLog('error', `Ошибка загрузки "${file.name}": ${(err as Error).message}`);
      }
    }

    setReports((prev) => [...prev, ...newReports]);
    if (template && newReports.length > 0) {
      setStep(2);
    }
    e.target.value = '';
  };

  // Remove a report
  const removeReport = (id: string) => {
    setReports((prev) => prev.filter((r) => r.id !== id));
  };

  // Toggle key column
  const toggleKeyColumn = (col: string) => {
    setKeyColumns((prev) =>
      prev.includes(col) ? prev.filter((c) => c !== col) : [...prev, col]
    );
  };

  // Toggle target column
  const toggleTargetColumn = (col: string) => {
    setTargetColumns((prev) =>
      prev.includes(col) ? prev.filter((c) => c !== col) : [...prev, col]
    );
  };

  // Check if columns are compatible between template and reports
  const getCommonColumns = (): string[] => {
    if (!template || reports.length === 0) return [];
    let common = [...template.data.headers];
    for (const report of reports) {
      common = common.filter((col) => report.data.headers.includes(col));
    }
    return common;
  };

  // Process matching
  const processMatching = () => {
    if (!template || reports.length === 0 || keyColumns.length === 0 || targetColumns.length === 0) {
      addLog('error', 'Не все параметры заданы для обработки');
      return;
    }

    setIsProcessing(true);
    clearLogs();
    addLog('info', 'Начало обработки...');
    addLog('info', `Ключевые столбцы: ${keyColumns.join(', ')}`);
    addLog('info', `Целевые столбцы: ${targetColumns.join(', ')}`);
    addLog('info', `Шаблон строк: ${template.data.rows.length}`);

    // Deep copy template rows
    const resultRows: Record<string, unknown>[] = template.data.rows.map((row) => ({ ...row }));

    let totalMatches = 0;
    let totalSkipped = 0;
    let totalFilled = 0;
    let totalNotFound = 0;

    for (const report of reports) {
      addLog('info', `--- Обработка отчёта: "${report.name}" (${report.data.rows.length} строк) ---`);

      for (let i = 0; i < resultRows.length; i++) {
        const templateRow = resultRows[i];

        // Build key from template row
        const templateKey = keyColumns.map((col) => String(templateRow[col] ?? '')).join('|||');

        // Find matching rows in report
        const matchingRows = report.data.rows.filter((reportRow) => {
          const reportKey = keyColumns.map((col) => String(reportRow[col] ?? '')).join('|||');
          return reportKey === templateKey;
        });

        if (matchingRows.length === 0) {
          totalNotFound++;
          continue;
        }

        if (matchingRows.length > 1) {
          totalSkipped++;
          const keyDisplay = keyColumns.map((col) => `${col}="${templateRow[col]}"`).join(', ');
          addLog(
            'warning',
            `Строка ${i + 1}: Найдено ${matchingRows.length} совпадений по ключу [${keyDisplay}]. Пропускаю.`
          );
          continue;
        }

        // Exactly one match - fill target columns
        const sourceRow = matchingRows[0];
        let filledCount = 0;
        for (const targetCol of targetColumns) {
          const value = sourceRow[targetCol];
          if (value !== undefined && value !== '') {
            templateRow[targetCol] = value;
            filledCount++;
          }
        }

        if (filledCount > 0) {
          totalMatches++;
          totalFilled += filledCount;
        }
      }

      addLog('success', `Отчёт "${report.name}" обработан`);
    }

    addLog('info', '=== ИТОГИ ===');
    addLog('success', `Успешно заполнено строк: ${totalMatches}`);
    addLog('info', `Заполнено ячеек: ${totalFilled}`);
    if (totalSkipped > 0) {
      addLog('warning', `Пропущено строк (множественные совпадения): ${totalSkipped}`);
    }
    if (totalNotFound > 0) {
      addLog('info', `Строк без совпадений в отчётах: ${totalNotFound}`);
    }

    setResultData({ headers: template.data.headers, rows: resultRows });
    setStep(4);
    setIsProcessing(false);
  };

  // Export result
  const exportResult = () => {
    if (!resultData) return;

    const ws = XLSX.utils.json_to_sheet(resultData.rows, { header: resultData.headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Результат');
    XLSX.writeFile(wb, 'result.xlsx');
    addLog('success', 'Файл result.xlsx экспортирован');
  };

  const commonColumns = getCommonColumns();

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white shadow-sm border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-4 py-4 sm:px-6 lg:px-8">
          <div className="flex items-center gap-3">
            <FileSpreadsheet className="w-8 h-8 text-emerald-600" />
            <div>
              <h1 className="text-2xl font-bold text-gray-900">Excel Merge Tool</h1>
              <p className="text-sm text-gray-500">
                Заполнение шаблона данными из отчётов по ключевым столбцам
              </p>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6 sm:px-6 lg:px-8">
        {/* Steps indicator */}
        <div className="mb-8 flex items-center justify-center gap-2">
          {[
            { num: 1, label: 'Загрузка файлов' },
            { num: 2, label: 'Настройка столбцов' },
            { num: 3, label: 'Обработка' },
            { num: 4, label: 'Результат' },
          ].map((s, idx) => (
            <div key={s.num} className="flex items-center gap-2">
              <div
                className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-medium ${
                  step >= s.num
                    ? 'bg-emerald-600 text-white'
                    : 'bg-gray-200 text-gray-500'
                }`}
              >
                {s.num}
              </div>
              <span
                className={`text-sm hidden sm:inline ${
                  step >= s.num ? 'text-emerald-700 font-medium' : 'text-gray-400'
                }`}
              >
                {s.label}
              </span>
              {idx < 3 && <div className="w-8 h-0.5 bg-gray-200" />}
            </div>
          ))}
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Left column - File uploads and settings */}
          <div className="lg:col-span-2 space-y-6">
            {/* Step 1: File Upload */}
            <section className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
              <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
                <Upload className="w-5 h-5 text-emerald-600" />
                Шаг 1: Загрузка файлов
              </h2>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {/* Template upload */}
                <div className="border-2 border-dashed border-gray-300 rounded-lg p-4 text-center hover:border-emerald-400 transition-colors">
                  <FileSpreadsheet className="w-10 h-10 mx-auto text-emerald-500 mb-2" />
                  <p className="text-sm font-medium text-gray-700 mb-1">Шаблон</p>
                  {template ? (
                    <div className="mt-2">
                      <p className="text-sm text-emerald-700 font-medium truncate">
                        {template.name}
                      </p>
                      <p className="text-xs text-gray-500">
                        {template.data.headers.length} столбцов, {template.data.rows.length} строк
                      </p>
                    </div>
                  ) : (
                    <p className="text-xs text-gray-400 mb-2">Выберите файл шаблона</p>
                  )}
                  <label className="mt-2 inline-block cursor-pointer bg-emerald-600 text-white text-sm px-3 py-1.5 rounded-lg hover:bg-emerald-700 transition-colors">
                    {template ? 'Заменить' : 'Выбрать файл'}
                    <input
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={handleTemplateUpload}
                      className="hidden"
                    />
                  </label>
                </div>

                {/* Reports upload */}
                <div className="border-2 border-dashed border-gray-300 rounded-lg p-4 text-center hover:border-blue-400 transition-colors">
                  <Upload className="w-10 h-10 mx-auto text-blue-500 mb-2" />
                  <p className="text-sm font-medium text-gray-700 mb-1">Отчёты</p>
                  {reports.length > 0 ? (
                    <div className="mt-2 space-y-1 max-h-32 overflow-y-auto">
                      {reports.map((r) => (
                        <div
                          key={r.id}
                          className="flex items-center justify-between bg-blue-50 rounded px-2 py-1"
                        >
                          <span className="text-xs text-blue-800 truncate flex-1 mr-1">
                            {r.name} ({r.data.rows.length} стр.)
                          </span>
                          <button
                            onClick={() => removeReport(r.id)}
                            className="text-red-400 hover:text-red-600"
                          >
                            <X className="w-3 h-3" />
                          </button>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p className="text-xs text-gray-400 mb-2">Выберите файлы отчётов</p>
                  )}
                  <label className="mt-2 inline-block cursor-pointer bg-blue-600 text-white text-sm px-3 py-1.5 rounded-lg hover:bg-blue-700 transition-colors">
                    {reports.length > 0 ? 'Добавить ещё' : 'Выбрать файлы'}
                    <input
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      multiple
                      onChange={handleReportsUpload}
                      className="hidden"
                    />
                  </label>
                </div>
              </div>
            </section>

            {/* Step 2: Column Configuration */}
            {(template && reports.length > 0) && (
              <section className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
                <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
                  <Settings className="w-5 h-5 text-emerald-600" />
                  Шаг 2: Настройка столбцов
                </h2>

                {commonColumns.length === 0 ? (
                  <div className="bg-red-50 border border-red-200 rounded-lg p-4 text-red-700 text-sm">
                    <AlertTriangle className="w-4 h-4 inline mr-1" />
                    Нет общих столбцов между шаблоном и отчётами. Проверьте структуру файлов.
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {/* Key columns */}
                    <div>
                      <h3 className="text-sm font-semibold text-gray-700 mb-2 flex items-center gap-1">
                        <span className="w-2 h-2 rounded-full bg-amber-500" />
                        Ключевые столбцы (для сопоставления)
                      </h3>
                      <p className="text-xs text-gray-500 mb-3">
                        Выберите столбцы, по которым будут искаться совпадения строк
                      </p>
                      <div className="border border-gray-200 rounded-lg max-h-64 overflow-y-auto">
                        {commonColumns.map((col) => (
                          <label
                            key={col}
                            className={`flex items-center gap-2 px-3 py-2 cursor-pointer hover:bg-amber-50 border-b border-gray-100 last:border-0 ${
                              keyColumns.includes(col) ? 'bg-amber-50' : ''
                            }`}
                          >
                            <input
                              type="checkbox"
                              checked={keyColumns.includes(col)}
                              onChange={() => toggleKeyColumn(col)}
                              className="rounded border-gray-300 text-amber-600 focus:ring-amber-500"
                            />
                            <span className="text-sm text-gray-700 truncate">{col}</span>
                          </label>
                        ))}
                      </div>
                      {keyColumns.length > 0 && (
                        <p className="mt-2 text-xs text-amber-700">
                          Выбрано: {keyColumns.length}
                        </p>
                      )}
                    </div>

                    {/* Target columns */}
                    <div>
                      <h3 className="text-sm font-semibold text-gray-700 mb-2 flex items-center gap-1">
                        <span className="w-2 h-2 rounded-full bg-emerald-500" />
                        Целевые столбцы (для заполнения)
                      </h3>
                      <p className="text-xs text-gray-500 mb-3">
                        Выберите столбцы, значения которых будут перенесены из отчётов в шаблон
                      </p>
                      <div className="border border-gray-200 rounded-lg max-h-64 overflow-y-auto">
                        {commonColumns
                          .filter((col) => !keyColumns.includes(col))
                          .map((col) => (
                            <label
                              key={col}
                              className={`flex items-center gap-2 px-3 py-2 cursor-pointer hover:bg-emerald-50 border-b border-gray-100 last:border-0 ${
                                targetColumns.includes(col) ? 'bg-emerald-50' : ''
                              }`}
                            >
                              <input
                                type="checkbox"
                                checked={targetColumns.includes(col)}
                                onChange={() => toggleTargetColumn(col)}
                                className="rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
                              />
                              <span className="text-sm text-gray-700 truncate">{col}</span>
                            </label>
                          ))}
                      </div>
                      {targetColumns.length > 0 && (
                        <p className="mt-2 text-xs text-emerald-700">
                          Выбрано: {targetColumns.length}
                        </p>
                      )}
                    </div>
                  </div>
                )}

                {/* Process button */}
                <div className="mt-6 flex items-center gap-4">
                  <button
                    onClick={processMatching}
                    disabled={
                      keyColumns.length === 0 ||
                      targetColumns.length === 0 ||
                      isProcessing ||
                      commonColumns.length === 0
                    }
                    className="flex items-center gap-2 bg-emerald-600 text-white px-6 py-2.5 rounded-lg font-medium hover:bg-emerald-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
                  >
                    <Play className="w-4 h-4" />
                    {isProcessing ? 'Обработка...' : 'Выполнить сопоставление'}
                  </button>

                  {keyColumns.length === 0 && (
                    <span className="text-xs text-gray-500">
                      ← Выберите хотя бы один ключевой столбец
                    </span>
                  )}
                  {keyColumns.length > 0 && targetColumns.length === 0 && (
                    <span className="text-xs text-gray-500">
                      ← Выберите хотя бы один целевой столбец
                    </span>
                  )}
                </div>
              </section>
            )}

            {/* Step 4: Result preview */}
            {resultData && (
              <section className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
                <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
                  <Download className="w-5 h-5 text-emerald-600" />
                  Результат
                </h2>

                <div className="flex items-center gap-4 mb-4">
                  <button
                    onClick={exportResult}
                    className="flex items-center gap-2 bg-emerald-600 text-white px-5 py-2.5 rounded-lg font-medium hover:bg-emerald-700 transition-colors"
                  >
                    <Download className="w-4 h-4" />
                    Скачать result.xlsx
                  </button>
                  <span className="text-sm text-gray-500">
                    {resultData.rows.length} строк, {resultData.headers.length} столбцов
                  </span>
                </div>

                {/* Preview table */}
                <div className="border border-gray-200 rounded-lg overflow-auto max-h-96">
                  <table className="min-w-full text-xs">
                    <thead className="bg-gray-50 sticky top-0">
                      <tr>
                        <th className="px-2 py-1.5 text-left font-medium text-gray-500 border-b">#</th>
                        {resultData.headers.slice(0, 10).map((h) => (
                          <th
                            key={h}
                            className={`px-2 py-1.5 text-left font-medium border-b ${
                              targetColumns.includes(h)
                                ? 'bg-emerald-50 text-emerald-700'
                                : keyColumns.includes(h)
                                ? 'bg-amber-50 text-amber-700'
                                : 'text-gray-600'
                            }`}
                          >
                            {h}
                          </th>
                        ))}
                        {resultData.headers.length > 10 && (
                          <th className="px-2 py-1.5 text-left font-medium text-gray-400 border-b">
                            +{resultData.headers.length - 10} ещё
                          </th>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {resultData.rows.slice(0, 20).map((row, i) => (
                        <tr key={i} className="hover:bg-gray-50">
                          <td className="px-2 py-1 text-gray-400 border-b">{i + 1}</td>
                          {resultData.headers.slice(0, 10).map((h) => (
                            <td
                              key={h}
                              className={`px-2 py-1 border-b truncate max-w-32 ${
                                targetColumns.includes(h) ? 'bg-emerald-50/50' : ''
                              }`}
                            >
                              {String(row[h] ?? '')}
                            </td>
                          ))}
                          {resultData.headers.length > 10 && (
                            <td className="px-2 py-1 text-gray-400 border-b">...</td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {resultData.rows.length > 20 && (
                  <p className="mt-2 text-xs text-gray-500 text-center">
                    Показано 20 из {resultData.rows.length} строк
                  </p>
                )}
              </section>
            )}
          </div>

          {/* Right column - Logs */}
          <div className="space-y-4">
            <section className="bg-white rounded-xl shadow-sm border border-gray-200 p-4 sticky top-4">
              <div className="flex items-center justify-between mb-3">
                <h2 className="text-sm font-semibold text-gray-700 flex items-center gap-2">
                  <Info className="w-4 h-4 text-gray-500" />
                  Журнал операций
                </h2>
                {logs.length > 0 && (
                  <button
                    onClick={clearLogs}
                    className="text-xs text-gray-400 hover:text-red-500 flex items-center gap-1"
                  >
                    <Trash2 className="w-3 h-3" />
                    Очистить
                  </button>
                )}
              </div>

              <div className="space-y-1.5 max-h-[70vh] overflow-y-auto">
                {logs.length === 0 ? (
                  <p className="text-xs text-gray-400 text-center py-8">
                    Журнал пуст. Загрузите файлы и запустите обработку.
                  </p>
                ) : (
                  logs.map((log, i) => (
                    <div
                      key={i}
                      className={`flex items-start gap-2 text-xs p-2 rounded ${
                        log.type === 'error'
                          ? 'bg-red-50 text-red-700'
                          : log.type === 'warning'
                          ? 'bg-amber-50 text-amber-700'
                          : log.type === 'success'
                          ? 'bg-emerald-50 text-emerald-700'
                          : 'bg-gray-50 text-gray-600'
                      }`}
                    >
                      {log.type === 'error' && <AlertTriangle className="w-3 h-3 mt-0.5 shrink-0" />}
                      {log.type === 'warning' && <AlertTriangle className="w-3 h-3 mt-0.5 shrink-0" />}
                      {log.type === 'success' && <CheckCircle className="w-3 h-3 mt-0.5 shrink-0" />}
                      {log.type === 'info' && <Info className="w-3 h-3 mt-0.5 shrink-0" />}
                      <span className="break-all">{log.message}</span>
                    </div>
                  ))
                )}
              </div>
            </section>

            {/* Info card */}
            <section className="bg-blue-50 rounded-xl border border-blue-200 p-4">
              <h3 className="text-sm font-semibold text-blue-800 mb-2">Как это работает</h3>
              <ol className="text-xs text-blue-700 space-y-1.5 list-decimal list-inside">
                <li>Загрузите файл-шаблон (который нужно заполнить)</li>
                <li>Загрузите один или несколько файлов-отчётов</li>
                <li>Выберите ключевые столбцы для сопоставления строк</li>
                <li>Выберите столбцы для переноса данных</li>
                <li>Нажмите "Выполнить сопоставление"</li>
              </ol>
              <div className="mt-3 p-2 bg-blue-100 rounded text-xs text-blue-800">
                <strong>Важно:</strong> Если по ключевым столбцам найдено более одного совпадения — строка будет пропущена с записью в журнал.
              </div>
            </section>
          </div>
        </div>
      </main>
    </div>
  );
}
