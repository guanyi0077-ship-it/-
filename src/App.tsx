/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { 
  Download, 
  Trash2, 
  Sparkles,
  FileCode,
  Play,
  Clock,
  AlertCircle,
  CheckCircle,
  Pencil,
  Square,
  ArrowUpRight,
  Eraser,
  Film,
  Image as ImageIcon,
  Sun,
  Moon,
  Folder,
  FileSpreadsheet,
  FileText,
  FileStack,
  Keyboard,
  Info
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { GoogleGenAI } from "@google/genai";
import ExcelJS from 'exceljs';
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  ImageRun, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  AlignmentType, 
  HeadingLevel,
  BorderStyle,
  VerticalAlign
} from 'docx';

declare global {
  interface Window {
    CSInterface: any;
  }
}

// Initialize Gemini for AI parsing
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

interface MarkerData {
  time: string;
  seconds: number;
  comment: string;
  category: 'visual' | 'audio' | 'edit' | 'technical' | 'general';
  color: number; // PR Marker Color Index: 0=Green, 1=Red, 2=Yellow, etc.
  screenshot?: string; // base64 image
  isConfirmed?: boolean;
}

interface VideoFile {
  id: string;
  file: File;
  url: string;
  name: string;
  size: string;
  resolution: string;
  markers: MarkerData[];
  isDetecting?: boolean;
  detectionProgress?: number;
}

const CATEGORY_COLORS: Record<MarkerData['category'], { label: string, color: string, prIndex: number }> = {
  visual: { label: '内容修改', color: '#3B82F6', prIndex: 4 }, // Blue
  audio: { label: '音频调整', color: '#10B981', prIndex: 0 },   // Green
  edit: { label: '剪辑建议', color: '#EF4444', prIndex: 1 },    // Red
  technical: { label: '技术问题', color: '#8B5CF6', prIndex: 3 }, // Purple
  general: { label: '常规批注', color: '#F59E0B', prIndex: 2 }, // Yellow
};

export default function App() {
  const [comment, setComment] = useState('');
  const [theme, setTheme] = useState<'light' | 'dark'>('dark');
  const [markers, setMarkers] = useState<MarkerData[]>([]);
  const [videoUrl, setVideoUrl] = useState<string | null>(null);
  const [videoList, setVideoList] = useState<VideoFile[]>([]);
  const [selectedVideoId, setSelectedVideoId] = useState<string | null>(null);
  const [currentTime, setCurrentTime] = useState(0);
  const [duration, setDuration] = useState(0);
  const [playbackRate, setPlaybackRate] = useState(1.0);
  const [hoveredMarker, setHoveredMarker] = useState<MarkerData | null>(null);
  const [selectedCategory, setSelectedCategory] = useState<MarkerData['category']>('general');
  const [successMessage, setSuccessMessage] = useState<string | null>(null);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editingText, setEditingText] = useState('');
  const [showShortcuts, setShowShortcuts] = useState(false);
  const [isDetecting, setIsDetecting] = useState(false);
  const [detectionProgress, setDetectionProgress] = useState(0);
  const stopDetectionRef = React.useRef(false);
  const [drawingMode, setDrawingMode] = useState<'none' | 'brush' | 'arrow' | 'rect'>('none');
  const [drawingColor, setDrawingColor] = useState('#F27D26');
  const [isDrawing, setIsDrawing] = useState(false);
  const [startPos, setStartPos] = useState({ x: 0, y: 0 });
  const [videoMetadata, setVideoMetadata] = useState<{
    name: string;
    size: string;
    resolution: string;
  } | null>(null);
  
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const drawingCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const tempCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const commentInputRef = React.useRef<HTMLTextAreaElement>(null);

  // Handle Video Upload
  const handleVideoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      const newVideos: VideoFile[] = [];
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        if (file.type.startsWith('video/')) {
          const url = URL.createObjectURL(file);
          const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
          const id = Math.random().toString(36).substr(2, 9);
          newVideos.push({
            id,
            file,
            url,
            name: file.name,
            size: `${sizeInMB} MB`,
            resolution: '正在加载...',
            markers: []
          });
        }
      }
      
      if (newVideos.length > 0) {
        setVideoList(prev => [...prev, ...newVideos]);
        // If no video is selected, select the first one from the new batch
        if (!selectedVideoId) {
          selectVideo(newVideos[0].id, [...videoList, ...newVideos]);
        }
        setSuccessMessage(`成功导入 ${newVideos.length} 个视频`);
        setTimeout(() => setSuccessMessage(null), 3000);
      }
    }
  };

  const selectVideo = (id: string, currentList?: VideoFile[]) => {
    const list = currentList || videoList;
    const video = list.find(v => v.id === id);
    if (video) {
      // Save current markers to the previously selected video
      if (selectedVideoId) {
        setVideoList(prev => prev.map(v => 
          v.id === selectedVideoId ? { ...v, markers: [...markers] } : v
        ));
      }

      setVideoUrl(video.url);
      setSelectedVideoId(id);
      setMarkers(video.markers);
      setVideoMetadata({
        name: video.name,
        size: video.size,
        resolution: video.resolution
      });
      setCurrentTime(0);
      if (videoRef.current) {
        videoRef.current.currentTime = 0;
      }
    }
  };

  // Sync Video Time
  const handleTimeUpdate = () => {
    if (videoRef.current) {
      setCurrentTime(videoRef.current.currentTime);
    }
  };

  // Auto-detect black frames when video changes
  React.useEffect(() => {
    if (videoUrl) {
      // Small delay to ensure everything is ready
      const timer = setTimeout(() => {
        detectBlackFrames();
      }, 1000);
      return () => clearTimeout(timer);
    }
  }, [videoUrl]);

  // Sync markers to videoList whenever they change
  React.useEffect(() => {
    if (selectedVideoId) {
      setVideoList(prev => prev.map(v => 
        v.id === selectedVideoId ? { ...v, markers: [...markers] } : v
      ));
    }
  }, [markers]);

  const handleLoadedMetadata = () => {
    if (videoRef.current && drawingCanvasRef.current && tempCanvasRef.current) {
      setDuration(videoRef.current.duration);
      
      const resolution = `${videoRef.current.videoWidth} x ${videoRef.current.videoHeight}`;
      
      // Update resolution in metadata
      if (videoMetadata) {
        setVideoMetadata({
          ...videoMetadata,
          resolution: resolution,
        });
      }

      // Update resolution in videoList
      if (selectedVideoId) {
        setVideoList(prev => prev.map(v => 
          v.id === selectedVideoId ? { ...v, resolution: resolution } : v
        ));
      }
      
      // Sync canvas size with video display size
      syncCanvasSize();
    }
  };

  const syncCanvasSize = () => {
    if (videoRef.current && drawingCanvasRef.current && tempCanvasRef.current) {
      const rect = videoRef.current.getBoundingClientRect();
      drawingCanvasRef.current.width = rect.width;
      drawingCanvasRef.current.height = rect.height;
      tempCanvasRef.current.width = rect.width;
      tempCanvasRef.current.height = rect.height;
    }
  };

  React.useEffect(() => {
    window.addEventListener('resize', syncCanvasSize);
    return () => window.removeEventListener('resize', syncCanvasSize);
  }, []);

  // Keyboard Shortcuts (Standard NLE Style)
  React.useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Don't trigger if user is typing in an input or textarea
      if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) {
        // Allow Escape to blur the input
        if (e.key === 'Escape') {
          (e.target as HTMLElement).blur();
        }
        return;
      }

      if (!videoRef.current) return;

      const video = videoRef.current;
      const frameTime = 1 / 30; // Assume 30fps for frame stepping

      switch (e.code) {
        case 'Space':
        case 'KeyK':
          e.preventDefault();
          if (video.paused) video.play();
          else video.pause();
          break;
        case 'ArrowLeft':
        case 'KeyJ':
          e.preventDefault();
          video.currentTime = Math.max(0, video.currentTime - 5);
          break;
        case 'ArrowRight':
        case 'KeyL':
          e.preventDefault();
          video.currentTime = Math.min(video.duration, video.currentTime + 5);
          break;
        case 'Comma': // Previous Frame
          e.preventDefault();
          video.pause();
          video.currentTime = Math.max(0, video.currentTime - frameTime);
          break;
        case 'Period': // Next Frame
          e.preventDefault();
          video.pause();
          video.currentTime = Math.min(video.duration, video.currentTime + frameTime);
          break;
        case 'KeyD': // Legacy Previous Frame
          e.preventDefault();
          video.pause();
          video.currentTime = Math.max(0, video.currentTime - frameTime);
          break;
        case 'KeyF': // Fullscreen (Standard)
          e.preventDefault();
          if (document.fullscreenElement) {
            document.exitFullscreen();
          } else {
            video.parentElement?.requestFullscreen();
          }
          break;
        case 'KeyM': // Focus Comment Input
          e.preventDefault();
          commentInputRef.current?.focus();
          break;
        case 'Digit1': setSelectedCategory('visual'); break;
        case 'Digit2': setSelectedCategory('audio'); break;
        case 'Digit3': setSelectedCategory('edit'); break;
        case 'Digit4': setSelectedCategory('technical'); break;
        case 'Digit5': setSelectedCategory('general'); break;
        case 'KeyC': // Speed Up
          e.preventDefault();
          const newRateC = Math.round((video.playbackRate + 0.1) * 10) / 10;
          video.playbackRate = newRateC;
          setPlaybackRate(newRateC);
          break;
        case 'KeyX': // Slow Down
          e.preventDefault();
          const newRateX = Math.max(0.1, Math.round((video.playbackRate - 0.1) * 10) / 10);
          video.playbackRate = newRateX;
          setPlaybackRate(newRateX);
          break;
        case 'KeyZ': // Reset Speed
          e.preventDefault();
          video.playbackRate = 1.0;
          setPlaybackRate(1.0);
          break;
        case 'ArrowUp':
          e.preventDefault();
          video.volume = Math.min(1, video.volume + 0.1);
          break;
        case 'ArrowDown':
          e.preventDefault();
          video.volume = Math.max(0, video.volume - 0.1);
          break;
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [videoUrl]);

  // Drawing Logic
  const startDrawing = (e: React.MouseEvent | React.TouchEvent) => {
    if (drawingMode === 'none') return;
    setIsDrawing(true);
    const rect = drawingCanvasRef.current?.getBoundingClientRect();
    if (!rect) return;
    
    const x = ('touches' in e) ? e.touches[0].clientX - rect.left : (e as React.MouseEvent).clientX - rect.left;
    const y = ('touches' in e) ? e.touches[0].clientY - rect.top : (e as React.MouseEvent).clientY - rect.top;
    
    setStartPos({ x, y });
    
    const ctx = drawingCanvasRef.current?.getContext('2d');
    if (ctx && drawingMode === 'brush') {
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.strokeStyle = drawingColor;
      ctx.lineWidth = 3;
      ctx.lineCap = 'round';
    }
  };

  const handleContextMenu = (e: React.MouseEvent) => {
    if (drawingMode !== 'none') {
      e.preventDefault();
      setDrawingMode('none');
    }
  };

  const draw = (e: React.MouseEvent | React.TouchEvent) => {
    if (!isDrawing || drawingMode === 'none') return;
    const rect = drawingCanvasRef.current?.getBoundingClientRect();
    if (!rect || !drawingCanvasRef.current || !tempCanvasRef.current) return;
    
    const x = ('touches' in e) ? e.touches[0].clientX - rect.left : (e as React.MouseEvent).clientX - rect.left;
    const y = ('touches' in e) ? e.touches[0].clientY - rect.top : (e as React.MouseEvent).clientY - rect.top;
    
    const ctx = drawingCanvasRef.current.getContext('2d');
    const tempCtx = tempCanvasRef.current.getContext('2d');
    if (!ctx || !tempCtx) return;

    if (drawingMode === 'brush') {
      ctx.lineTo(x, y);
      ctx.stroke();
    } else {
      tempCtx.clearRect(0, 0, tempCanvasRef.current.width, tempCanvasRef.current.height);
      tempCtx.strokeStyle = drawingColor;
      tempCtx.lineWidth = 3;
      if (drawingMode === 'rect') {
        tempCtx.strokeRect(startPos.x, startPos.y, x - startPos.x, y - startPos.y);
      } else if (drawingMode === 'arrow') {
        drawArrow(tempCtx, startPos.x, startPos.y, x, y, drawingColor);
      }
    }
  };

  const stopDrawing = (e: React.MouseEvent | React.TouchEvent) => {
    if (!isDrawing) return;
    
    const rect = drawingCanvasRef.current?.getBoundingClientRect();
    if (!rect || !drawingCanvasRef.current || !tempCanvasRef.current) return;
    
    const x = ('touches' in e) ? (e as any).changedTouches[0].clientX - rect.left : (e as React.MouseEvent).clientX - rect.left;
    const y = ('touches' in e) ? (e as any).changedTouches[0].clientY - rect.top : (e as React.MouseEvent).clientY - rect.top;
    
    const ctx = drawingCanvasRef.current.getContext('2d');
    const tempCtx = tempCanvasRef.current.getContext('2d');
    if (ctx && tempCtx) {
      if (drawingMode === 'rect') {
        ctx.strokeStyle = drawingColor;
        ctx.lineWidth = 3;
        ctx.strokeRect(startPos.x, startPos.y, x - startPos.x, y - startPos.y);
      } else if (drawingMode === 'arrow') {
        drawArrow(ctx, startPos.x, startPos.y, x, y, drawingColor);
      }
      tempCtx.clearRect(0, 0, tempCanvasRef.current.width, tempCanvasRef.current.height);
    }
    
    setIsDrawing(false);
  };

  const drawArrow = (ctx: CanvasRenderingContext2D, fromX: number, fromY: number, toX: number, toY: number, color: string) => {
    const headLength = 15;
    const angle = Math.atan2(toY - fromY, toX - fromX);
    
    ctx.strokeStyle = color;
    ctx.lineWidth = 3;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    
    ctx.beginPath();
    ctx.moveTo(fromX, fromY);
    ctx.lineTo(toX, toY);
    ctx.stroke();
    
    ctx.beginPath();
    ctx.moveTo(toX, toY);
    ctx.lineTo(toX - headLength * Math.cos(angle - Math.PI / 6), toY - headLength * Math.sin(angle - Math.PI / 6));
    ctx.moveTo(toX, toY);
    ctx.lineTo(toX - headLength * Math.cos(angle + Math.PI / 6), toY - headLength * Math.sin(angle + Math.PI / 6));
    ctx.stroke();
  };

  const clearCanvas = () => {
    const ctx = drawingCanvasRef.current?.getContext('2d');
    if (ctx && drawingCanvasRef.current) {
      ctx.clearRect(0, 0, drawingCanvasRef.current.width, drawingCanvasRef.current.height);
    }
  };

  const seekTo = (seconds: number) => {
    if (videoRef.current) {
      videoRef.current.currentTime = seconds;
      videoRef.current.pause(); // Pause at the point of review
    }
  };

  // Add Review Comment
  const handleAddReview = () => {
    if (!videoRef.current) return;

    // Check for duplicates (same time and same comment)
    const isDuplicate = markers.some(m => 
      m.seconds === currentTime && 
      m.comment.trim() === comment.trim() &&
      m.category === selectedCategory
    );

    if (isDuplicate) {
      setSuccessMessage('该时间点已存在相同的批注');
      setTimeout(() => setSuccessMessage(null), 2000);
      return;
    }

    // Capture screenshot
    const captureCanvas = document.createElement('canvas');
    captureCanvas.width = videoRef.current.videoWidth;
    captureCanvas.height = videoRef.current.videoHeight;
    const ctx = captureCanvas.getContext('2d');
    
    if (ctx) {
      // Draw video frame
      ctx.drawImage(videoRef.current, 0, 0, captureCanvas.width, captureCanvas.height);
      
      // Draw annotations if any
      if (drawingCanvasRef.current) {
        ctx.drawImage(drawingCanvasRef.current, 0, 0, captureCanvas.width, captureCanvas.height);
      }
    }

    const screenshot = captureCanvas.toDataURL('image/jpeg', 0.8);

    const newMarker: MarkerData = {
      time: new Date(currentTime * 1000).toISOString().substr(11, 8),
      seconds: currentTime,
      comment: comment,
      category: selectedCategory,
      color: CATEGORY_COLORS[selectedCategory].prIndex,
      screenshot: screenshot
    };

    setMarkers(prev => [...prev, newMarker].sort((a, b) => a.seconds - b.seconds));
    setComment('');
    clearCanvas();
  };

  const deleteMarker = (index: number) => {
    setMarkers(prev => prev.filter((_, i) => i !== index));
    if (editingIndex === index) {
      setEditingIndex(null);
    }
  };

  const startEditing = (index: number, text: string) => {
    setEditingIndex(index);
    setEditingText(text);
  };

  const saveEdit = (index: number) => {
    setMarkers(prev => {
      const newMarkers = [...prev];
      newMarkers[index] = { ...newMarkers[index], comment: editingText };
      return newMarkers;
    });
    setEditingIndex(null);
  };

  const cancelEdit = () => {
    setEditingIndex(null);
    setEditingText('');
  };

  const confirmMarker = (index: number) => {
    setMarkers(prev => {
      const newMarkers = [...prev];
      newMarkers[index] = { ...newMarkers[index], isConfirmed: true };
      return newMarkers;
    });
  };

  // Black Frame Detection Logic
  const detectBlackFrames = async () => {
    if (!videoUrl || !videoRef.current) return;
    
    if (isDetecting) {
      stopDetectionRef.current = true;
      return;
    }

    setIsDetecting(true);
    stopDetectionRef.current = false;
    setDetectionProgress(0);
    setSuccessMessage('正在检测黑帧...');

    const video = document.createElement('video');
    video.src = videoUrl;
    video.crossOrigin = 'anonymous';
    video.muted = true;
    
    try {
      await new Promise((resolve, reject) => {
        video.onloadedmetadata = resolve;
        video.onerror = () => reject(new Error('视频加载失败'));
        // Timeout after 10s
        setTimeout(() => reject(new Error('加载超时')), 10000);
      });

      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      if (!ctx) throw new Error('无法创建画布上下文');

      const duration = video.duration;
      const interval = 0.5; // Check every 0.5 seconds
      const threshold = 12; // Brightness threshold (0-255)
      const detectedMarkers: MarkerData[] = [];
      
      canvas.width = 160; // Low res for speed
      canvas.height = 90;

      for (let time = 0; time < duration; time += interval) {
        if (stopDetectionRef.current) break;

        video.currentTime = time;
        await new Promise((resolve) => {
          video.onseeked = resolve;
        });

        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const pixels = imageData.data;
        
        let totalBrightness = 0;
        for (let i = 0; i < pixels.length; i += 4) {
          totalBrightness += (pixels[i] + pixels[i+1] + pixels[i+2]) / 3;
        }
        
        const avgBrightness = totalBrightness / (pixels.length / 4);
        
        if (avgBrightness < threshold) {
          const isNear = detectedMarkers.some(m => Math.abs(m.seconds - time) < 2);
          if (!isNear) {
            const screenshot = canvas.toDataURL('image/jpeg', 0.5);
            detectedMarkers.push({
              time: new Date(time * 1000).toISOString().substr(11, 8),
              seconds: time,
              comment: '检测到疑似黑帧，请确认',
              category: 'technical',
              color: CATEGORY_COLORS['technical'].prIndex,
              screenshot: screenshot,
              isConfirmed: false
            });
          }
        }
        
        setDetectionProgress(Math.round((time / duration) * 100));
      }

      if (stopDetectionRef.current) {
        setSuccessMessage('检测已停止');
      } else if (detectedMarkers.length > 0) {
        setMarkers(prev => [...prev, ...detectedMarkers].sort((a, b) => a.seconds - b.seconds));
        setSuccessMessage(`检测完成，共发现 ${detectedMarkers.length} 处疑似黑帧`);
      } else {
        setSuccessMessage('检测完成，未发现黑帧');
      }
    } catch (err) {
      console.error('Detection error:', err);
      setSuccessMessage(err instanceof Error ? err.message : '检测出错');
    } finally {
      setIsDetecting(false);
      setDetectionProgress(0);
      stopDetectionRef.current = false;
      setTimeout(() => setSuccessMessage(null), 3000);
    }
  };

  // Export Logic
  const exportReviewComments = async () => {
    if (markers.length === 0) return;
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('审核报告');

    // Add Video Info Section
    if (videoMetadata) {
      worksheet.mergeCells('A1:D1');
      const titleCell = worksheet.getCell('A1');
      titleCell.value = `视频审核报告 - ${videoMetadata.name}`;
      titleCell.font = { bold: true, size: 16 };
      titleCell.alignment = { horizontal: 'center' };

      worksheet.addRow(['视频名称', videoMetadata.name]);
      worksheet.addRow(['文件大小', videoMetadata.size]);
      worksheet.addRow(['分辨率', videoMetadata.resolution]);
      worksheet.addRow(['导出时间', new Date().toLocaleString()]);
      worksheet.addRow([]); // Empty row for spacing
    }

    // Define columns starting from the next available row
    const headerRowIndex = videoMetadata ? 7 : 1;
    
    worksheet.getRow(headerRowIndex).values = ['时间戳', '类别', '审核意见', '截图预览'];
    worksheet.columns = [
      { key: 'time', width: 15 },
      { key: 'category', width: 15 },
      { key: 'comment', width: 40 },
      { key: 'screenshot', width: 60 }
    ];

    // Style header
    const headerRow = worksheet.getRow(headerRowIndex);
    headerRow.font = { bold: true, size: 12 };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.height = 30;

    // Add markers
    for (let i = 0; i < markers.length; i++) {
      const m = markers[i];
      const currentRowIndex = headerRowIndex + i + 1;
      const row = worksheet.addRow({
        time: m.time,
        category: CATEGORY_COLORS[m.category].label,
        comment: m.comment
      });

      // Set row height for image
      row.height = 180;
      row.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

      if (m.screenshot) {
        try {
          // More robust base64 stripping using regex
          const base64Data = m.screenshot.replace(/^data:image\/\w+;base64,/, "");

          const imageId = workbook.addImage({
            base64: base64Data,
            extension: 'jpeg',
          });

          worksheet.addImage(imageId, {
            tl: { col: 3, row: currentRowIndex - 1 },
            ext: { width: 400, height: 225 }
          });
        } catch (err) {
          console.error('Failed to add image to Excel:', err);
        }
      }
    }

    // Generate and download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const fileName = videoMetadata 
      ? `审核报告_${videoMetadata.name.split('.')[0]}_${new Date().toISOString().split('T')[0]}.xlsx`
      : `视频审核报告_${new Date().toLocaleDateString()}.xlsx`;
    a.download = fileName;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportBatchReport = async () => {
    if (videoList.length === 0) return;
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('视频列表汇总');
    
    worksheet.columns = [
      { header: '视频名称', key: 'name', width: 40 },
      { header: '文件大小', key: 'size', width: 15 },
      { header: '分辨率', key: 'resolution', width: 20 },
      { header: '批注数量', key: 'markerCount', width: 15 },
      { header: '技术问题(黑帧)', key: 'technicalCount', width: 20 },
    ];
    
    // Sync current markers to videoList before export
    const updatedList = videoList.map(v => 
      v.id === selectedVideoId ? { ...v, markers: [...markers] } : v
    );

    updatedList.forEach(v => {
      const technicalCount = v.markers.filter(m => m.category === 'technical').length;
      worksheet.addRow({
        name: v.name,
        size: v.size,
        resolution: v.resolution,
        markerCount: v.markers.length,
        technicalCount: technicalCount
      });
    });
    
    // Styling
    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF27D26' }
    };

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `批量审核汇总_${new Date().toLocaleDateString()}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportAllMarkersToExcel = async () => {
    if (videoList.length === 0) return;
    
    // Sync current markers to videoList before export
    const updatedList = videoList.map(v => 
      v.id === selectedVideoId ? { ...v, markers: [...markers] } : v
    );

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('所有视频批注汇总');
    
    worksheet.columns = [
      { header: '视频名称', key: 'videoName', width: 30 },
      { header: '时间点', key: 'time', width: 15 },
      { header: '类别', key: 'category', width: 15 },
      { header: '修改意见', key: 'comment', width: 50 },
      { header: '截图', key: 'screenshot', width: 60 },
    ];

    // Header styling
    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF27D26' }
    };

    let currentRowIndex = 2;
    for (const v of updatedList) {
      if (v.markers.length === 0) continue;
      
      for (const m of v.markers) {
        const row = worksheet.addRow({
          videoName: v.name,
          time: m.time,
          category: CATEGORY_COLORS[m.category].label,
          comment: m.comment
        });

        row.height = 180;
        row.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

        if (m.screenshot) {
          try {
            const base64Data = m.screenshot.replace(/^data:image\/\w+;base64,/, "");
            const imageId = workbook.addImage({
              base64: base64Data,
              extension: 'jpeg',
            });

            worksheet.addImage(imageId, {
              tl: { col: 4, row: currentRowIndex - 1 },
              ext: { width: 400, height: 225 }
            });
          } catch (err) {
            console.error('Failed to add image to Excel:', err);
          }
        }
        currentRowIndex++;
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `全量视频批注汇总_${new Date().toLocaleDateString()}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportAllMarkersToWord = async () => {
    if (videoList.length === 0) return;

    // Sync current markers to videoList before export
    const updatedList = videoList.map(v => 
      v.id === selectedVideoId ? { ...v, markers: [...markers] } : v
    );

    const docSections = [];

    for (const v of updatedList) {
      if (v.markers.length === 0) continue;

      const videoHeader = new Paragraph({
        text: `视频: ${v.name}`,
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 200 },
      });

      const tableRows = [
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "时间点", style: "HeaderStyle" })], width: { size: 15, type: WidthType.PERCENTAGE } }),
            new TableCell({ children: [new Paragraph({ text: "类别", style: "HeaderStyle" })], width: { size: 15, type: WidthType.PERCENTAGE } }),
            new TableCell({ children: [new Paragraph({ text: "修改意见", style: "HeaderStyle" })], width: { size: 30, type: WidthType.PERCENTAGE } }),
            new TableCell({ children: [new Paragraph({ text: "截图", style: "HeaderStyle" })], width: { size: 40, type: WidthType.PERCENTAGE } }),
          ],
        }),
      ];

      for (const m of v.markers) {
        const screenshotCellChildren = [];
        if (m.screenshot) {
          try {
            const base64Data = m.screenshot.replace(/^data:image\/\w+;base64,/, "");
            const binary = atob(base64Data);
            const bytes = new Uint8Array(binary.length);
            for (let i = 0; i < binary.length; i++) {
              bytes[i] = binary.charCodeAt(i);
            }

            screenshotCellChildren.push(
              new Paragraph({
                children: [
                  new ImageRun({
                    data: bytes,
                    transformation: { width: 300, height: 168 },
                    type: "jpg",
                  }),
                ],
              })
            );
          } catch (err) {
            console.error('Failed to add image to Word:', err);
          }
        }

        tableRows.push(
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph(m.time)], verticalAlign: VerticalAlign.CENTER }),
              new TableCell({ children: [new Paragraph(CATEGORY_COLORS[m.category].label)], verticalAlign: VerticalAlign.CENTER }),
              new TableCell({ children: [new Paragraph(m.comment)], verticalAlign: VerticalAlign.CENTER }),
              new TableCell({ children: screenshotCellChildren, verticalAlign: VerticalAlign.CENTER }),
            ],
          })
        );
      }

      const table = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tableRows,
      });

      docSections.push({
        children: [videoHeader, table],
      });
    }

    if (docSections.length === 0) {
      setSuccessMessage('没有批注可以导出');
      setTimeout(() => setSuccessMessage(null), 2000);
      return;
    }

    const doc = new Document({
      styles: {
        paragraphStyles: [
          {
            id: "HeaderStyle",
            name: "Header Style",
            basedOn: "Normal",
            next: "Normal",
            run: { bold: true, color: "F27D26" },
          },
        ],
      },
      sections: docSections,
    });

    const buffer = await Packer.toBlob(doc);
    const url = URL.createObjectURL(buffer);
    const a = document.createElement('a');
    a.href = url;
    a.download = `全量视频批注汇总_${new Date().toLocaleDateString()}.docx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className={`min-h-screen transition-colors duration-300 ${theme === 'dark' ? 'bg-[#0A0A0A] text-[#E0E0E0]' : 'bg-[#F5F5F3] text-[#1A1A1A]'} font-sans selection:bg-[#F27D26]`}>
      <div className="max-w-[1600px] mx-auto p-4 md:p-6 h-screen flex flex-col">
        
        {/* Header */}
        <header className="flex items-center justify-between mb-6 shrink-0 relative">
          <AnimatePresence>
            {successMessage && (
              <motion.div
                initial={{ opacity: 0, y: -20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -20 }}
                className="absolute left-1/2 -translate-x-1/2 top-0 bg-[#F27D26] text-white px-4 py-2 rounded-full text-xs font-bold shadow-lg z-50"
              >
                {successMessage}
              </motion.div>
            )}
          </AnimatePresence>
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-[#F27D26] rounded-xl flex items-center justify-center shadow-lg shadow-[#F27D26]/20">
              <Play className="text-white fill-current" size={20} />
            </div>
            <div>
              <h1 className={`text-xl font-bold tracking-tight ${theme === 'dark' ? 'text-white' : 'text-[#1A1A1A]'}`}>ReviewFlow <span className="text-[#F27D26] text-xs font-normal ml-2">Studio v1.0</span></h1>
              <p className={`text-[9px] uppercase tracking-[0.2em] font-mono ${theme === 'dark' ? 'opacity-30' : 'opacity-50'}`}>Professional Video Review Platform</p>
            </div>
          </div>
          
          <div className="flex items-center gap-4">
            <button 
              onClick={() => setShowShortcuts(true)}
              className={`p-2 rounded-lg transition-all ${theme === 'dark' ? 'hover:bg-white/5 text-white/40' : 'hover:bg-black/5 text-black/40'}`}
              title="键盘快捷键"
            >
              <Keyboard size={18} />
            </button>
            {videoUrl && (
              <button 
                onClick={detectBlackFrames}
                disabled={isDetecting}
                className={`flex items-center gap-2 px-3 py-2 rounded-lg text-[10px] font-bold uppercase transition-all border ${isDetecting ? 'opacity-50 cursor-not-allowed' : 'hover:scale-105 active:scale-95'} ${theme === 'dark' ? 'bg-purple-500/10 border-purple-500/30 text-purple-400' : 'bg-purple-500/5 border-purple-500/20 text-purple-600'}`}
              >
                <Sparkles size={14} className={isDetecting ? 'animate-spin' : ''} />
                {isDetecting ? `检测中 ${detectionProgress}%` : '黑帧检测'}
              </button>
            )}
            {videoList.length > 1 && (
              <div className="flex items-center gap-2">
                <button 
                  onClick={exportBatchReport}
                  className={`flex items-center gap-2 px-3 py-2 rounded-lg text-[10px] font-bold uppercase transition-all border ${theme === 'dark' ? 'bg-green-500/10 border-green-500/30 text-green-400' : 'bg-green-500/5 border-green-500/20 text-green-600'}`}
                  title="导出视频列表汇总"
                >
                  <FileSpreadsheet size={14} />
                  汇总表
                </button>
                <button 
                  onClick={exportAllMarkersToExcel}
                  className={`flex items-center gap-2 px-3 py-2 rounded-lg text-[10px] font-bold uppercase transition-all border ${theme === 'dark' ? 'bg-blue-500/10 border-blue-500/30 text-blue-400' : 'bg-blue-500/5 border-blue-500/20 text-blue-600'}`}
                  title="导出所有视频的详细批注意见"
                >
                  <FileStack size={14} />
                  全量Excel
                </button>
                <button 
                  onClick={exportAllMarkersToWord}
                  className={`flex items-center gap-2 px-3 py-2 rounded-lg text-[10px] font-bold uppercase transition-all border ${theme === 'dark' ? 'bg-amber-500/10 border-amber-500/30 text-amber-400' : 'bg-amber-500/5 border-amber-500/20 text-amber-600'}`}
                  title="导出所有视频的详细批注意见 (Word)"
                >
                  <FileText size={14} />
                  全量Word
                </button>
              </div>
            )}
            <button 
              onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
              className={`p-2 rounded-lg transition-all border ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 border-white/10 text-white' : 'bg-black/5 hover:bg-black/10 border-black/10 text-black'}`}
            >
              {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
            </button>
            <label className={`cursor-pointer px-4 py-2 text-[10px] font-bold uppercase transition-all rounded-lg flex items-center gap-2 border ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 border-white/10 text-white' : 'bg-black/5 hover:bg-black/10 border-black/10 text-black'}`}>
              <Download size={14} />
              导入视频
              <input type="file" accept="video/*" multiple onChange={handleVideoUpload} className="hidden" />
            </label>
            <label className={`cursor-pointer px-4 py-2 text-[10px] font-bold uppercase transition-all rounded-lg flex items-center gap-2 border ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 border-white/10 text-white' : 'bg-black/5 hover:bg-black/10 border-black/10 text-black'}`}>
              <Folder size={14} />
              导入文件夹
              <input 
                type="file" 
                webkitdirectory="" 
                directory="" 
                multiple 
                onChange={handleVideoUpload} 
                className="hidden" 
              />
            </label>
          </div>
        </header>

        <main className="flex-1 grid grid-cols-1 lg:grid-cols-12 gap-6 min-h-0">
          
          {/* Left: Player & Input Section (Expanded) */}
          <div className="lg:col-span-9 flex flex-col gap-6 min-h-0 overflow-y-auto custom-scrollbar pr-2">
            {/* Video Container */}
            <div className={`bg-black rounded-none overflow-hidden shadow-2xl border relative aspect-video flex items-center justify-center group shrink-0 ${theme === 'dark' ? 'border-white/5' : 'border-black/5'}`}>
              {videoUrl ? (
                <>
                  <video 
                    ref={videoRef}
                    src={videoUrl}
                    onTimeUpdate={handleTimeUpdate}
                    onLoadedMetadata={handleLoadedMetadata}
                    className="max-w-full max-h-full"
                    controls
                  />
                  <canvas
                    ref={drawingCanvasRef}
                    onMouseDown={startDrawing}
                    onMouseMove={draw}
                    onMouseUp={stopDrawing}
                    onMouseLeave={stopDrawing}
                    onContextMenu={handleContextMenu}
                    onTouchStart={startDrawing}
                    onTouchMove={draw}
                    onTouchEnd={stopDrawing}
                    className={`absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 ${drawingMode === 'none' ? 'pointer-events-none' : 'pointer-events-auto cursor-crosshair'}`}
                  />
                  <canvas
                    ref={tempCanvasRef}
                    className={`absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 pointer-events-none ${drawingMode === 'none' ? 'hidden' : ''}`}
                  />

                  {/* Floating Drawing Toolbar */}
                  <div className="absolute top-4 left-1/2 -translate-x-1/2 flex items-center gap-3 bg-black/80 backdrop-blur-xl border border-white/10 p-2 rounded-2xl opacity-0 group-hover:opacity-100 transition-all duration-300 z-30 shadow-2xl scale-90 group-hover:scale-100">
                    <div className="flex items-center gap-1.5 px-1">
                      {['#F27D26', '#EF4444', '#22C55E', '#3B82F6', '#EAB308'].map(color => (
                        <button
                          key={color}
                          onClick={() => setDrawingColor(color)}
                          className={`w-4 h-4 rounded-full border transition-all ${drawingColor === color ? 'border-white scale-125' : 'border-transparent opacity-50 hover:opacity-100'}`}
                          style={{ backgroundColor: color }}
                        />
                      ))}
                    </div>
                    <div className="w-px h-4 bg-white/10" />
                    <div className="flex items-center gap-1">
                      <button 
                        onClick={() => setDrawingMode(drawingMode === 'brush' ? 'none' : 'brush')}
                        className={`p-2 rounded-xl transition-all ${drawingMode === 'brush' ? 'bg-[#F27D26] text-white' : 'hover:bg-white/10 text-white/40'}`}
                        title="画笔"
                      >
                        <Pencil size={14} />
                      </button>
                      <button 
                        onClick={() => setDrawingMode(drawingMode === 'rect' ? 'none' : 'rect')}
                        className={`p-2 rounded-xl transition-all ${drawingMode === 'rect' ? 'bg-[#F27D26] text-white' : 'hover:bg-white/10 text-white/40'}`}
                        title="框选"
                      >
                        <Square size={14} />
                      </button>
                      <button 
                        onClick={() => setDrawingMode(drawingMode === 'arrow' ? 'none' : 'arrow')}
                        className={`p-2 rounded-xl transition-all ${drawingMode === 'arrow' ? 'bg-[#F27D26] text-white' : 'hover:bg-white/10 text-white/40'}`}
                        title="箭头"
                      >
                        <ArrowUpRight size={14} />
                      </button>
                    </div>
                    <div className="w-px h-4 bg-white/10" />
                    <button 
                      onClick={clearCanvas}
                      className="p-2 rounded-xl hover:bg-red-500/20 text-red-500 transition-all"
                      title="清除批注"
                    >
                      <Eraser size={14} />
                    </button>
                  </div>
                </>
              ) : (
                <div className="flex flex-col items-center opacity-20">
                  <Play size={80} strokeWidth={1} />
                  <p className="mt-6 text-xs uppercase tracking-[0.3em]">等待视频载入...</p>
                </div>
              )}
              
              {/* Timeline Markers Overlay */}
              <div className="absolute bottom-[60px] left-10 right-10 h-1 pointer-events-none">
                {markers.map((m, i) => (
                  <div 
                    key={i} 
                    className="absolute top-0 w-3 h-3 -translate-x-1/2 -translate-y-1/2 rounded-full border-2 border-black shadow-xl pointer-events-auto cursor-pointer hover:scale-150 transition-transform"
                    style={{ 
                      left: `${(m.seconds / duration) * 100}%`,
                      backgroundColor: CATEGORY_COLORS[m.category].color 
                    }}
                    onClick={() => seekTo(m.seconds)}
                    onMouseEnter={() => setHoveredMarker(m)}
                    onMouseLeave={() => setHoveredMarker(null)}
                  />
                ))}
              </div>
            </div>

            {/* Video Info Bar (Below Video) */}
            {videoUrl && (
              <div className={`flex items-center justify-between px-4 py-3 border rounded-2xl shrink-0 ${theme === 'dark' ? 'bg-[#111] border-white/5' : 'bg-white border-black/5'}`}>
                <div className="flex items-center gap-3">
                  <div className={`${theme === 'dark' ? 'bg-[#F27D26]/10' : 'bg-[#F27D26]/5'} p-2 rounded-lg`}>
                    <Film size={16} className="text-[#F27D26]" />
                  </div>
                  <div>
                    <h2 className={`text-xs font-bold tracking-tight ${theme === 'dark' ? 'text-white/90' : 'text-[#1A1A1A]'}`}>
                      {videoMetadata?.name || '未命名视频'}
                    </h2>
                    <div className="flex items-center gap-3 mt-0.5">
                      <p className={`text-[8px] font-bold uppercase tracking-wider ${theme === 'dark' ? 'text-white/30' : 'text-black/30'}`}>正在审阅</p>
                      <div className={`w-px h-2 ${theme === 'dark' ? 'bg-white/10' : 'bg-black/10'}`} />
                      <span className={`text-[8px] font-mono ${
                        videoMetadata?.resolution && 
                        videoMetadata.resolution !== '1920 x 1080' && 
                        videoMetadata.resolution !== '正在加载...' 
                          ? 'text-red-500 font-bold' 
                          : (theme === 'dark' ? 'text-white/40' : 'text-black/40')
                      }`}>
                        {videoMetadata?.resolution || '--'}
                      </span>
                      <div className={`w-px h-2 ${theme === 'dark' ? 'bg-white/10' : 'bg-black/10'}`} />
                      <span className={`text-[8px] font-mono ${theme === 'dark' ? 'text-white/40' : 'text-black/40'}`}>{videoMetadata?.size || '--'}</span>
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-4">
                  <span className={`text-[10px] font-mono text-[#F27D26] px-2 py-1 rounded-lg ${theme === 'dark' ? 'bg-[#F27D26]/10' : 'bg-[#F27D26]/5'}`}>
                    {new Date(currentTime * 1000).toISOString().substr(11, 8)} / {new Date(duration * 1000).toISOString().substr(11, 8)}
                  </span>
                </div>
              </div>
            )}

            {/* Review Input Section (Below Video) */}
            <div className={`flex flex-col border rounded-3xl overflow-hidden shadow-2xl transition-colors duration-300 shrink-0 ${theme === 'dark' ? 'bg-[#111] border-white/5' : 'bg-white border-black/5'}`}>
              <div className={`p-4 border-b ${theme === 'dark' ? 'bg-white/[0.02] border-white/5' : 'bg-black/[0.02] border-black/5'}`}>
                <h3 className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2">
                  <Sparkles size={14} className="text-[#F27D26]" />
                  添加审核意见
                </h3>
              </div>
              <div className="p-4 space-y-4">
                <div className="flex flex-wrap gap-1.5">
                  {(Object.keys(CATEGORY_COLORS) as Array<keyof typeof CATEGORY_COLORS>).map(cat => (
                    <button
                      key={cat}
                      onClick={() => setSelectedCategory(cat)}
                      className={`px-2 py-1 rounded-lg text-[8px] font-bold uppercase transition-all ${
                        selectedCategory === cat 
                        ? (theme === 'dark' ? 'bg-white text-black' : 'bg-black text-white')
                        : (theme === 'dark' ? 'bg-white/5 text-white/40 hover:bg-white/10' : 'bg-black/5 text-black/40 hover:bg-black/10')
                      }`}
                    >
                      {CATEGORY_COLORS[cat].label}
                    </button>
                  ))}
                </div>
                <div className="relative">
                  <textarea
                    ref={commentInputRef}
                    value={comment}
                    onChange={(e) => setComment(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && e.ctrlKey) {
                        e.preventDefault();
                        handleAddReview();
                      }
                    }}
                    placeholder="在此输入您的审核意见... (Ctrl+Enter 发送)"
                    className={`w-full border rounded-2xl p-4 text-xs focus:outline-none focus:border-[#F27D26]/40 transition-all resize-none h-24 ${theme === 'dark' ? 'bg-black/40 border-white/5 text-white' : 'bg-black/5 border-black/5 text-black'}`}
                  />
                  <button
                    onClick={handleAddReview}
                    disabled={!videoUrl}
                    className="absolute bottom-4 right-4 bg-[#F27D26] text-white p-2.5 rounded-xl hover:scale-105 active:scale-95 transition-all disabled:opacity-20 shadow-lg shadow-[#F27D26]/20"
                  >
                    <Sparkles size={16} />
                  </button>
                </div>
              </div>
            </div>
          </div>

          {/* Right: Sidebar Section (Review List & Video List) */}
          <div className="lg:col-span-3 flex flex-col gap-6 min-h-0">
            {/* Review List Section */}
            <div className={`flex-1 flex flex-col border rounded-3xl overflow-hidden shadow-2xl min-h-0 transition-colors duration-300 ${theme === 'dark' ? 'bg-[#111] border-white/5' : 'bg-white border-black/5'}`}>
              <div className={`p-4 border-b flex items-center justify-between ${theme === 'dark' ? 'bg-white/[0.02] border-white/5' : 'bg-black/[0.02] border-black/5'}`}>
                <h3 className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2">
                  <Clock size={14} className="text-[#F27D26]" />
                  批注列表 ({markers.length})
                </h3>
                <button onClick={() => setMarkers([])} className={`p-1.5 rounded-xl transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/20' : 'hover:bg-black/5 text-black/20'} hover:text-red-500`}>
                  <Trash2 size={12} />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto custom-scrollbar p-3 space-y-3">
                {markers.length === 0 ? (
                  <div className="h-full flex flex-col items-center justify-center opacity-10 text-center px-6">
                    <FileCode size={32} strokeWidth={1} />
                    <p className="mt-3 text-[8px] uppercase tracking-widest leading-relaxed">暂无批注</p>
                  </div>
                ) : (
                  markers.map((m, i) => (
                    <motion.div
                      initial={{ opacity: 0, x: 20 }}
                      animate={{ opacity: 1, x: 0 }}
                      key={i}
                      onClick={() => seekTo(m.seconds)}
                      className={`group relative border rounded-2xl p-3 cursor-pointer transition-all ${theme === 'dark' ? 'bg-white/[0.03] hover:bg-white/[0.06] border-white/5' : 'bg-black/[0.03] hover:bg-black/[0.06] border-black/5'}`}
                    >
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[9px] font-mono text-[#F27D26] font-bold">{m.time}</span>
                        <span 
                          className="text-[7px] font-bold uppercase px-1.5 py-0.5 rounded"
                          style={{ backgroundColor: `${CATEGORY_COLORS[m.category].color}20`, color: CATEGORY_COLORS[m.category].color }}
                        >
                          {CATEGORY_COLORS[m.category].label}
                        </span>
                      </div>
                      <p className={`text-[11px] leading-relaxed line-clamp-2 ${theme === 'dark' ? 'text-white/80' : 'text-black/80'}`}>
                        {m.comment || <span className="italic opacity-30">(仅截图)</span>}
                      </p>
                      <div className="absolute top-2 right-2 flex gap-1 opacity-0 group-hover:opacity-100 transition-all">
                        <button onClick={(e) => { e.stopPropagation(); deleteMarker(i); }} className="p-1 hover:bg-red-500/20 text-red-500 rounded-lg transition-all"><Trash2 size={10} /></button>
                      </div>
                    </motion.div>
                  ))
                )}
              </div>

              <div className="p-3 border-t">
                <button 
                  onClick={exportReviewComments}
                  disabled={markers.length === 0}
                  className="w-full py-2.5 rounded-2xl bg-[#F27D26] text-white text-[9px] font-bold uppercase hover:scale-[1.02] active:scale-[0.98] transition-all disabled:opacity-20 flex items-center justify-center gap-2 shadow-lg shadow-[#F27D26]/20"
                >
                  <Download size={14} />
                  导出批注
                </button>
              </div>
            </div>

            {/* Video List Section (Bottom of Sidebar) */}
            {videoList.length > 0 && (
              <div className={`h-[250px] flex flex-col border rounded-3xl overflow-hidden shadow-2xl transition-colors duration-300 ${theme === 'dark' ? 'bg-[#111] border-white/5' : 'bg-white border-black/5'}`}>
                <div className={`p-4 border-b flex items-center justify-between ${theme === 'dark' ? 'bg-white/[0.02] border-white/5' : 'bg-black/[0.02] border-black/5'}`}>
                  <h3 className="text-[10px] font-bold uppercase tracking-widest flex items-center gap-2">
                    <Film size={14} className="text-[#F27D26]" />
                    视频列表 ({videoList.length})
                  </h3>
                  <button 
                    onClick={() => { setVideoList([]); setVideoUrl(null); setSelectedVideoId(null); setMarkers([]); setVideoMetadata(null); }}
                    className={`p-1 rounded-lg transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/20' : 'hover:bg-black/5 text-black/20'} hover:text-red-500`}
                  >
                    <Trash2 size={12} />
                  </button>
                </div>
                <div className="flex-1 overflow-y-auto custom-scrollbar p-2 space-y-1">
                  {videoList.map((v) => (
                    <div key={v.id} className="relative group">
                      <button
                        onClick={() => selectVideo(v.id)}
                        className={`w-full text-left p-2 rounded-xl transition-all relative ${
                          selectedVideoId === v.id 
                          ? (theme === 'dark' ? 'bg-[#F27D26]/20 text-[#F27D26]' : 'bg-[#F27D26]/10 text-[#F27D26]')
                          : (theme === 'dark' ? 'hover:bg-white/5 text-white/40' : 'hover:bg-black/5 text-black/60')
                        }`}
                      >
                        <span className="text-[9px] font-bold truncate block pr-6" title={v.name}>{v.name}</span>
                      </button>
                      <button
                        onClick={(e) => { e.stopPropagation(); setVideoList(prev => prev.filter(item => item.id !== v.id)); if (selectedVideoId === v.id) { setVideoUrl(null); setSelectedVideoId(null); setMarkers([]); setVideoMetadata(null); } }}
                        className="absolute right-1.5 top-1/2 -translate-y-1/2 p-1 opacity-0 group-hover:opacity-100 transition-all hover:text-red-500 text-white/20"
                      >
                        <Trash2 size={10} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </main>

        {/* Hover Tooltip for Timeline */}
        <AnimatePresence>
          {hoveredMarker && (
            <motion.div 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 10 }}
              className={`fixed bottom-32 left-1/2 -translate-x-1/2 border px-4 py-2 rounded-lg shadow-2xl z-50 pointer-events-none ${theme === 'dark' ? 'bg-[#1A1A1A] border-white/10 text-white' : 'bg-white border-black/10 text-black'}`}
            >
              <p className="text-[10px] text-[#F27D26] font-mono mb-1">{hoveredMarker.time}</p>
              <p className="text-xs font-medium">{hoveredMarker.comment || '(仅截图批注)'}</p>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
      {/* Keyboard Shortcuts Modal */}
      <AnimatePresence>
        {showShortcuts && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowShortcuts(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className={`relative w-full max-w-md border rounded-3xl p-6 shadow-2xl ${theme === 'dark' ? 'bg-[#111] border-white/10 text-white' : 'bg-white border-black/10 text-black'}`}
            >
              <div className="flex items-center justify-between mb-6">
                <h3 className="text-sm font-bold uppercase tracking-widest flex items-center gap-2">
                  <Keyboard size={18} className="text-[#F27D26]" />
                  键盘快捷键
                </h3>
                <button 
                  onClick={() => setShowShortcuts(false)}
                  className={`p-1.5 rounded-xl transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/20' : 'hover:bg-black/5 text-black/20'}`}
                >
                  <Trash2 size={14} className="rotate-45" />
                </button>
              </div>

              <div className="space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <p className="text-[10px] font-bold opacity-30 uppercase tracking-wider mb-2">播放控制</p>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">播放/暂停</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">Space / K</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">后退 5s</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">J / ←</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">前进 5s</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">L / →</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">上一帧</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">,</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">下一帧</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">.</kbd>
                    </div>
                  </div>
                  <div className="space-y-2">
                    <p className="text-[10px] font-bold opacity-30 uppercase tracking-wider mb-2">审核功能</p>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">聚焦输入框</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">M</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">发送批注</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">Ctrl+Enter</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">选择分类</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">1 - 5</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">全屏播放</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">F</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">退出输入</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">Esc</kbd>
                    </div>
                  </div>
                </div>
                
                <div className="pt-4 border-t border-white/5">
                  <p className="text-[10px] font-bold opacity-30 uppercase tracking-wider mb-2">速度与音量</p>
                  <div className="grid grid-cols-2 gap-x-4 gap-y-2 text-xs">
                    <div className="flex justify-between items-center">
                      <span className="opacity-50">加速/减速</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">C / X</kbd>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="opacity-50">重置速度</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">Z</kbd>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="opacity-50">音量调节</span>
                      <kbd className="px-1.5 py-0.5 rounded bg-white/5 border border-white/10 font-mono text-[10px]">↑ / ↓</kbd>
                    </div>
                  </div>
                </div>
              </div>

              <div className="mt-8 flex items-center gap-2 p-3 rounded-2xl bg-[#F27D26]/10 border border-[#F27D26]/20">
                <Info size={14} className="text-[#F27D26] shrink-0" />
                <p className="text-[9px] text-[#F27D26] font-medium leading-relaxed">
                  提示：在输入框内时，快捷键将自动禁用，以便您正常输入。按 Esc 可快速退出输入状态。
                </p>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

