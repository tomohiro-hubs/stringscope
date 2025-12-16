import React, { useState, useEffect, useRef } from 'react';
import { Card, CardHeader, CardTitle, CardContent } from './components/ui/Card';
import { Button } from './components/ui/Button';
import { processFiles } from './utils/excelProcessor';
import { saveAs } from 'file-saver';
import { Upload, Clock, FileSpreadsheet, AlertTriangle, CheckCircle, Download, Loader2 } from 'lucide-react';

function App() {
  const [masterData, setMasterData] = useState(null);
  const [isLoadingMaster, setIsLoadingMaster] = useState(true);
  const [masterError, setMasterError] = useState(null);

  const [timeRange, setTimeRange] = useState({ start: '09:00', end: '18:00' });
  const [files, setFiles] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(null);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);

  const fileInputRef = useRef(null);

  useEffect(() => {
    fetch('data.json')
      .then(res => {
        if (!res.ok) throw new Error('マスタファイルが見つかりません');
        return res.json();
      })
      .then(data => {
        setMasterData(data);
        setIsLoadingMaster(false);
      })
      .catch(err => {
        console.error(err);
        setMasterError(err.message);
        setIsLoadingMaster(false);
      });
  }, []);

  const handleFileChange = (e) => {
    if (e.target.files) {
      setFiles(Array.from(e.target.files));
      setResult(null); // Reset previous result
      setError(null);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    if (e.dataTransfer.files) {
        const droppedFiles = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.xlsx'));
        if (droppedFiles.length > 0) {
            setFiles(droppedFiles);
            setResult(null);
            setError(null);
        }
    }
  };

  const handleRun = async () => {
    if (!masterData) return;
    setIsProcessing(true);
    setProgress({ phase: 'init', message: '処理を開始します...' });
    setError(null);

    try {
      const { buffer, stats, unknownPCSList } = await processFiles(
        files, 
        timeRange, 
        masterData, 
        (p) => {
            let msg = '';
            if (p.phase === 'reading') msg = `ファイル読込中 (${p.current}/${p.total}): ${p.filename}`;
            if (p.phase === 'processing') msg = `データ解析中... 行: ${p.current}`;
            setProgress({ ...p, message: msg });
        }
      );

      setResult({ buffer, stats, unknownPCSList });
    } catch (err) {
      console.error(err);
      setError(err.message || '予期せぬエラーが発生しました');
    } finally {
      setIsProcessing(false);
      setProgress(null);
    }
  };

  const downloadResult = () => {
    if (!result) return;
    const now = new Date();
    const dateStr = now.toISOString().slice(0,10).replace(/-/g, '');
    const timeStr = timeRange.start.replace(':','') + '-' + timeRange.end.replace(':','');
    const fileName = `merged_0A_highlight_${dateStr}_${timeStr}.xlsx`;
    
    const blob = new Blob([result.buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, fileName);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans pb-10">
      {/* Header */}
      <header className="bg-slate-900 text-white p-4 shadow-md">
        <div className="container mx-auto flex items-center gap-2">
            <FileSpreadsheet className="h-6 w-6 text-blue-400" />
            <h1 className="text-xl font-bold">PCSストリング電流解析ツール</h1>
        </div>
      </header>

      <main className="container mx-auto p-4 space-y-6 max-w-3xl">
        
        {/* Status Messages */}
        {isLoadingMaster && (
            <div className="bg-blue-50 border-l-4 border-blue-500 p-4 rounded text-blue-700 flex items-center">
                <Loader2 className="animate-spin mr-2 h-5 w-5" />
                PCSマスタデータを読み込み中...
            </div>
        )}
        {masterError && (
            <div className="bg-red-50 border-l-4 border-red-500 p-4 rounded text-red-700 flex items-center">
                <AlertTriangle className="mr-2 h-5 w-5" />
                <div>
                    <p className="font-bold">マスタ読込エラー</p>
                    <p>{masterError}</p>
                    <p className="text-sm mt-1">public/data.json が存在するか確認してください。</p>
                </div>
            </div>
        )}

        {/* Input Card */}
        <Card>
            <CardHeader>
                <CardTitle className="text-lg flex items-center gap-2">
                    <CheckCircle className="h-5 w-5 text-green-600" />
                    解析条件設定
                </CardTitle>
            </CardHeader>
            <CardContent className="space-y-6">
                
                {/* Time Range */}
                <div className="space-y-2">
                    <label className="text-sm font-medium text-slate-700 flex items-center gap-1">
                        <Clock className="h-4 w-4" /> 判定時間帯 (開始 - 終了)
                    </label>
                    <div className="flex items-center gap-4">
                        <input 
                            type="time" 
                            className="border border-slate-300 rounded px-3 py-2 w-full focus:outline-none focus:ring-2 focus:ring-blue-500"
                            value={timeRange.start}
                            onChange={e => setTimeRange({...timeRange, start: e.target.value})}
                        />
                        <span className="text-slate-400">～</span>
                        <input 
                            type="time" 
                            className="border border-slate-300 rounded px-3 py-2 w-full focus:outline-none focus:ring-2 focus:ring-blue-500"
                            value={timeRange.end}
                            onChange={e => setTimeRange({...timeRange, end: e.target.value})}
                        />
                    </div>
                    <p className="text-xs text-slate-500">
                        ※この時間帯に含まれるデータのみ0Aチェックを行います。終了時間が開始時間より前の場合、日跨ぎとみなします。
                    </p>
                </div>

                {/* File Upload */}
                <div className="space-y-2">
                    <label className="text-sm font-medium text-slate-700 flex items-center gap-1">
                        <Upload className="h-4 w-4" /> ログファイル (Excel .xlsx)
                    </label>
                    
                    <div 
                        className="border-2 border-dashed border-slate-300 rounded-lg p-8 text-center hover:bg-slate-50 transition-colors cursor-pointer"
                        onDragOver={e => e.preventDefault()}
                        onDrop={handleDrop}
                        onClick={() => fileInputRef.current?.click()}
                    >
                        <input 
                            type="file" 
                            multiple 
                            accept=".xlsx" 
                            className="hidden" 
                            ref={fileInputRef}
                            onChange={handleFileChange}
                        />
                        <FileSpreadsheet className="h-10 w-10 text-slate-400 mx-auto mb-2" />
                        <p className="text-slate-600 font-medium">クリックしてファイルを選択</p>
                        <p className="text-sm text-slate-400 mt-1">またはここにドラッグ＆ドロップ</p>
                    </div>

                    {files.length > 0 && (
                        <div className="bg-slate-100 p-3 rounded text-sm text-slate-700">
                            <strong>{files.length}</strong> ファイル選択中:
                            <ul className="list-disc list-inside mt-1 pl-2 text-xs text-slate-500 max-h-24 overflow-y-auto">
                                {files.map((f, i) => <li key={i}>{f.name} ({Math.round(f.size/1024)} KB)</li>)}
                            </ul>
                        </div>
                    )}
                </div>

                <Button 
                    className="w-full h-12 text-lg" 
                    disabled={!masterData || files.length === 0 || isProcessing}
                    onClick={handleRun}
                >
                    {isProcessing ? (
                        <span className="flex items-center gap-2">
                            <Loader2 className="animate-spin h-5 w-5" /> 処理中...
                        </span>
                    ) : '実行する'}
                </Button>

                {isProcessing && progress && (
                    <div className="space-y-1">
                        <div className="flex justify-between text-xs text-slate-500">
                            <span>{progress.message}</span>
                        </div>
                        <div className="h-2 bg-slate-200 rounded-full overflow-hidden">
                            <div className="h-full bg-blue-500 animate-pulse w-full"></div>
                        </div>
                    </div>
                )}
            </CardContent>
        </Card>

        {/* Error Display */}
        {error && (
            <div className="bg-red-50 border border-red-200 text-red-700 p-4 rounded-lg flex items-start gap-3">
                <AlertTriangle className="h-5 w-5 mt-0.5 flex-shrink-0" />
                <div>
                    <h4 className="font-bold">エラーが発生しました</h4>
                    <p>{error}</p>
                </div>
            </div>
        )}

        {/* Result Card */}
        {result && (
            <Card className="border-green-200 bg-green-50/50">
                <CardHeader>
                    <CardTitle className="text-green-800 flex items-center gap-2">
                        <CheckCircle className="h-6 w-6" /> 処理完了
                    </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                    <div className="grid grid-cols-2 gap-4 text-sm bg-white p-4 rounded border border-green-100">
                        <div>
                            <span className="text-slate-500 block">処理ファイル数</span>
                            <span className="font-bold text-lg">{result.stats.filesProcessed}</span>
                        </div>
                        <div>
                            <span className="text-slate-500 block">出力行数</span>
                            <span className="font-bold text-lg">{result.stats.totalRows}</span>
                        </div>
                        <div>
                            <span className="text-slate-500 block">判定対象行数</span>
                            <span className="font-bold text-lg">{result.stats.targetRows}</span>
                        </div>
                        <div>
                            <span className="text-slate-500 block">赤ハイライト(0A)</span>
                            <span className="font-bold text-lg text-red-600">{result.stats.highlightedCells} <span className="text-xs font-normal text-slate-400">セル</span></span>
                        </div>
                    </div>

                    {result.unknownPCSList.length > 0 && (
                        <div className="bg-yellow-50 border border-yellow-200 p-3 rounded text-sm text-yellow-800">
                            <strong className="flex items-center gap-1">
                                <AlertTriangle className="h-4 w-4" />
                                マスタ未登録のPCS ({result.stats.unknownPCS}件)
                            </strong>
                            <p className="text-xs mt-1 text-yellow-700">以下のPCSは回路数が不明なため、全PV(1-8)をチェックしました:</p>
                            <ul className="list-disc list-inside mt-1 text-xs max-h-20 overflow-y-auto">
                                {result.unknownPCSList.map(pcs => <li key={pcs}>{pcs}</li>)}
                            </ul>
                        </div>
                    )}

                    <Button 
                        onClick={downloadResult} 
                        className="w-full h-14 text-lg bg-green-600 hover:bg-green-700 shadow-lg shadow-green-200"
                    >
                        <Download className="mr-2 h-6 w-6" /> 結果ファイルをダウンロード
                    </Button>
                </CardContent>
            </Card>
        )}

      </main>

      <footer className="text-center text-slate-400 text-sm mt-10">
        <p>PCS 0A Detection Tool v1.0.0</p>
      </footer>
    </div>
  );
}

export default App;
