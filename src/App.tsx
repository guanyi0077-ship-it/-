/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
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
  Info,
  ClipboardList,
  ChevronDown,
  X,
  History,
  Plus,
  FileUp,
  Quote
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import * as MP4Box from 'mp4box';
import { 
  Volume2,
  VolumeX,
  Maximize2,
  Minimize2,
  Settings,
  Activity
} from 'lucide-react';
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

// PR Marker Color Index: 0=Green, 1=Red, 2=Yellow, etc.
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
  duration?: string;
  markers: MarkerData[];
  isDetecting?: boolean;
  detectionProgress?: number;
  fileHandle?: any;
}

const CATEGORY_COLORS: Record<MarkerData['category'], { label: string, color: string, prIndex: number }> = {
  visual: { label: '内容修改', color: '#3B82F6', prIndex: 4 }, // Blue
  audio: { label: '音频调整', color: '#10B981', prIndex: 0 },   // Green
  edit: { label: '剪辑建议', color: '#EF4444', prIndex: 1 },    // Red
  technical: { label: '技术问题', color: '#8B5CF6', prIndex: 3 }, // Purple
  general: { label: '常规批注', color: '#F59E0B', prIndex: 2 }, // Yellow
};

const DEFAULT_QUICK_REPLIES: Record<string, string[]> = {
  visual: [
    '画面构图不佳，建议居中。',
    '画面过暗，建议提高亮度。',
    '画面出现闪烁。',
    '建议增加特写镜头。',
    '颜色饱和度过高。'
  ],
  audio: [
    '音频背景噪音过大。',
    '音频左右声道不平衡。',
    '音频音量突降，请检查。'
  ],
  edit: [
    '建议在此处剪开。',
    '奏过慢，建议压缩。',
    '此处转场过于生硬。'
  ],
  technical: [
    '编码格式不符合规范。',
    '码率过低，画面出现马赛克。'
  ],
  general: [
    '字幕内容有误，请核对。',
    '演员表演略显生硬。',
    '灯光阴影处理不当。',
    '此处视频素材重复。'
  ]
};

// --- Audio Visualizer Component ---
const AudioVisualizer = ({ 
  theme, 
  analyserL, 
  analyserR,
  onImbalance,
}: { 
  theme: 'light' | 'dark',
  analyserL: AnalyserNode | null,
  analyserR: AnalyserNode | null,
  onImbalance?: (imbalanced: boolean) => void;
}) => {
  const meterLRef = React.useRef<HTMLDivElement>(null);
  const meterRRef = React.useRef<HTMLDivElement>(null);
  const valueLRef = React.useRef<HTMLSpanElement>(null);
  const valueRRef = React.useRef<HTMLSpanElement>(null);
  const requestRef = React.useRef<number>();
  const imbalanceCounterRef = React.useRef<number>(0);
  const onImbalanceRef = React.useRef(onImbalance);
  const [isLocalImbalanced, setIsLocalImbalanced] = React.useState(false);

  React.useEffect(() => {
    onImbalanceRef.current = onImbalance;
  }, [onImbalance]);

  // Update Volume Meters
  const animate = () => {
    let volL = 0;
    let volR = 0;

    if (analyserL && meterLRef.current && valueLRef.current) {
      const dataArray = new Uint8Array(analyserL.frequencyBinCount);
      analyserL.getByteFrequencyData(dataArray);
      const average = dataArray.reduce((a, b) => a + b) / dataArray.length;
      volL = Math.min(100, (average / 128) * 100);
      const db = Math.round(20 * Math.log10(average / 255 || 0.0001));
      
      meterLRef.current.style.height = `${volL}%`;
      meterLRef.current.style.backgroundColor = volL > 90 ? '#EF4444' : volL > 70 ? '#EAB308' : '#10B981';
      valueLRef.current.innerText = `${db}dB`;
    }
    if (analyserR && meterRRef.current && valueRRef.current) {
      const dataArray = new Uint8Array(analyserR.frequencyBinCount);
      analyserR.getByteFrequencyData(dataArray);
      const average = dataArray.reduce((a, b) => a + b) / dataArray.length;
      volR = Math.min(100, (average / 128) * 100);
      const db = Math.round(20 * Math.log10(average / 255 || 0.0001));

      meterRRef.current.style.height = `${volR}%`;
      meterRRef.current.style.backgroundColor = volR > 90 ? '#EF4444' : volR > 70 ? '#EAB308' : '#10B981';
      valueRRef.current.innerText = `${db}dB`;
    }

    // Imbalance detection logic
    // Extremely sensitive threshold for testing/visibility (15% difference)
    if (volL > 2 || volR > 2) {
      const diff = Math.abs(volL - volR);
      if (diff > 15) { 
        imbalanceCounterRef.current++;
      } else {
        imbalanceCounterRef.current = Math.max(0, imbalanceCounterRef.current - 2);
      }
    } else {
      imbalanceCounterRef.current = Math.max(0, imbalanceCounterRef.current - 1);
    }

    // Sustained for ~0.5 seconds
    if (imbalanceCounterRef.current > 30) {
      if (!isLocalImbalanced) {
        setIsLocalImbalanced(true);
        onImbalanceRef.current?.(true);
      }
    } else if (imbalanceCounterRef.current === 0) {
      if (isLocalImbalanced) {
        setIsLocalImbalanced(false);
        onImbalanceRef.current?.(false);
      }
    }

    requestRef.current = requestAnimationFrame(animate);
  };

  React.useEffect(() => {
    requestRef.current = requestAnimationFrame(animate);
    return () => {
      if (requestRef.current) cancelAnimationFrame(requestRef.current);
    };
  }, [analyserL, analyserR]);

  const dbScale = [0, -6, -12, -18, -24, -36, -48];

  return (
    <div className={`flex flex-col h-full border rounded-2xl overflow-hidden relative transition-all duration-300 ${isLocalImbalanced ? 'border-red-500 shadow-[0_0_15px_rgba(239,68,68,0.3)]' : (theme === 'dark' ? 'bg-black border-white/5' : 'bg-white border-black/5')}`}>
      {/* Imbalance Overlay */}
      <AnimatePresence>
        {isLocalImbalanced && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="absolute inset-0 bg-red-500/10 pointer-events-none z-20 animate-pulse"
          />
        )}
      </AnimatePresence>

      {/* Real-time DB Values at the Top */}
      <div className="flex justify-around px-1 py-1 bg-black/60 border-b border-white/5 shrink-0 relative z-30">
        <div className="flex flex-col items-center">
          <span className="text-[6px] font-bold text-white/20 uppercase tracking-tighter">L</span>
          <span ref={valueLRef} className="text-[8px] font-mono font-bold text-[#10B981]">-INF</span>
        </div>
        <div className="flex flex-col items-center">
          <span className="text-[6px] font-bold text-white/20 uppercase tracking-tighter">R</span>
          <span ref={valueRRef} className="text-[8px] font-mono font-bold text-[#3B82F6]">-INF</span>
        </div>
      </div>

      <div className="flex flex-1 min-h-0 p-1.5 gap-1.5 relative z-30">
        {/* Meters Container */}
        <div className="flex-1 flex gap-1 bg-black/40 rounded-lg p-1 relative">
          {/* Scale Lines & Labels (Overlay style) */}
          <div className="absolute inset-x-0 inset-y-1 flex flex-col justify-between pointer-events-none z-10">
            {dbScale.map(db => (
              <div key={db} className="flex items-center gap-1 w-full opacity-30">
                <div className={`h-px flex-1 ${db === 0 ? 'bg-red-500' : 'bg-white/20'}`} />
                <span className={`text-[5px] font-mono w-4 text-right ${db === 0 ? 'text-red-500' : 'text-white/40'}`}>{db}</span>
              </div>
            ))}
          </div>

          {/* Left Meter */}
          <div className="flex-1 bg-white/5 rounded-sm relative overflow-hidden flex flex-col justify-end">
            <div ref={meterLRef} className="w-full transition-all duration-75 ease-out" style={{ height: '0%' }} />
          </div>

          {/* Right Meter */}
          <div className="flex-1 bg-white/5 rounded-sm relative overflow-hidden flex flex-col justify-end">
            <div ref={meterRRef} className="w-full transition-all duration-75 ease-out" style={{ height: '0%' }} />
          </div>
        </div>
      </div>
      
      {/* Footer Activity Icon */}
      <div className="py-1 bg-black/80 border-t border-white/5 shrink-0 flex items-center justify-center relative z-30">
        {isLocalImbalanced ? (
          <AlertCircle size={10} className="text-red-500 animate-bounce" />
        ) : (
          <Activity size={10} className="text-[#10B981] opacity-50" />
        )}
      </div>
    </div>
  );
};

// --- Memoized UI Components for Performance ---

const QuickReplyItem = React.memo(({ 
  reply, 
  index, 
  tab, 
  theme, 
  onRemove, 
  onDoubleClick 
}: { 
  reply: string, 
  index: number, 
  tab: string, 
  theme: string, 
  onRemove: (tab: string, i: number) => void,
  onDoubleClick: (reply: string) => void
}) => (
  <motion.div
    initial={{ opacity: 0, x: -10 }}
    animate={{ opacity: 1, x: 0 }}
    onDoubleClick={() => onDoubleClick(reply)}
    className={`group relative border rounded-xl p-2 cursor-pointer transition-all select-none ${theme === 'dark' ? 'bg-white/[0.03] hover:bg-white/[0.06] border-white/5' : 'bg-white hover:bg-gray-50 border-gray-100 shadow-sm'}`}
    title="双击发送批注"
  >
    <p className={`text-[10px] leading-relaxed line-clamp-3 ${theme === 'dark' ? 'text-white/80' : 'text-gray-700 font-medium'}`}>
      {reply}
    </p>
    <button 
      onClick={(e) => { e.stopPropagation(); onRemove(tab, index); }} 
      className="absolute -top-1 -right-1 p-0.5 bg-red-500 text-white rounded-full opacity-0 group-hover:opacity-100 transition-all scale-75"
    >
      <X size={8} />
    </button>
  </motion.div>
));

const MarkerItem = React.memo(({ 
  marker, 
  index, 
  theme, 
  onSeek, 
  onDelete 
}: { 
  marker: MarkerData, 
  index: number, 
  theme: string, 
  onSeek: (s: number) => void,
  onDelete: (i: number) => void
}) => (
  <motion.div
    initial={{ opacity: 0, x: 20 }}
    animate={{ opacity: 1, x: 0 }}
    onClick={() => onSeek(marker.seconds)}
    className={`group relative border rounded-xl p-2 cursor-pointer transition-all ${theme === 'dark' ? 'bg-white/[0.03] hover:bg-white/[0.06] border-white/5' : 'bg-white hover:bg-gray-50 border-gray-100 shadow-sm'}`}
  >
    <div className="flex items-center justify-between mb-1">
      <span className="text-[8px] font-mono text-[#F27D26] font-bold">{marker.time}</span>
      <span 
        className="text-[6px] font-bold uppercase px-1 py-0.5 rounded"
        style={{ backgroundColor: `${CATEGORY_COLORS[marker.category].color}20`, color: CATEGORY_COLORS[marker.category].color }}
      >
        {CATEGORY_COLORS[marker.category].label}
      </span>
    </div>
    <p className={`text-[10px] leading-relaxed line-clamp-2 ${theme === 'dark' ? 'text-white/80' : 'text-gray-700 font-medium'}`}>
      {marker.comment || <span className="italic opacity-30">(仅截图)</span>}
    </p>
    <div className="absolute top-1 right-1 flex gap-1 opacity-0 group-hover:opacity-100 transition-all">
      <button onClick={(e) => { e.stopPropagation(); onDelete(index); }} className="p-1 hover:bg-red-500/20 text-red-500 rounded-lg transition-all"><Trash2 size={8} /></button>
    </div>
  </motion.div>
));

const VideoListItem = React.memo(({ 
  video, 
  isSelected, 
  theme, 
  onSelect, 
  onDelete 
}: { 
  video: VideoFile, 
  isSelected: boolean, 
  theme: string, 
  onSelect: (id: string) => void,
  onDelete: (id: string) => void
}) => (
  <div className="relative group">
    <button
      onClick={() => onSelect(video.id)}
      className={`w-full text-left p-1.5 rounded-lg transition-all relative ${
        isSelected 
        ? (theme === 'dark' ? 'bg-[#F27D26]/20 text-[#F27D26]' : 'bg-[#F27D26]/10 text-[#F27D26]')
        : (theme === 'dark' ? 'hover:bg-white/5 text-white/40' : 'hover:bg-gray-50 text-gray-600')
      }`}
    >
      <span className="text-[8px] font-bold truncate block pr-6" title={video.name}>{video.name}</span>
    </button>
    <button
      onClick={(e) => { e.stopPropagation(); onDelete(video.id); }}
      className={`absolute right-1 top-1/2 -translate-y-1/2 p-1 opacity-0 group-hover:opacity-100 transition-all hover:text-red-500 ${theme === 'dark' ? 'text-white/20' : 'text-black/20'}`}
    >
      <Trash2 size={8} />
    </button>
  </div>
));

const VideoControls = React.memo(({ 
  videoRef, 
  duration, 
  markers, 
  onSeek, 
  onSetHoveredMarker,
  isDetecting,
  detectionProgress,
  theme
}: { 
  videoRef: React.RefObject<HTMLVideoElement>, 
  duration: number, 
  markers: MarkerData[], 
  onSeek: (s: number) => void,
  onSetHoveredMarker: (m: MarkerData | null) => void,
  isDetecting: boolean,
  detectionProgress: number,
  theme: string
}) => {
  const [localTime, setLocalTime] = useState(0);
  const [isPaused, setIsPaused] = useState(true);
  const [isMuted, setIsMuted] = useState(false);
  const requestRef = useRef<number>();

  const updateLocalState = () => {
    if (videoRef.current) {
      setLocalTime(videoRef.current.currentTime);
      setIsPaused(videoRef.current.paused);
      setIsMuted(videoRef.current.muted);
    }
    requestRef.current = requestAnimationFrame(updateLocalState);
  };

  useEffect(() => {
    requestRef.current = requestAnimationFrame(updateLocalState);
    return () => {
      if (requestRef.current) cancelAnimationFrame(requestRef.current);
    };
  }, []);

  const formatTime = (s: number) => {
    return new Date(s * 1000).toISOString().substr(11, 8);
  };

  return (
    <div className="absolute bottom-0 left-0 right-0 bg-gradient-to-t from-black/80 via-black/20 to-transparent flex flex-col z-30 pointer-events-auto transition-all duration-300 group-hover:from-black/95 group-hover:via-black/60">
      {/* Progress Bar */}
      <div 
        className="relative h-1.5 w-full bg-white/10 cursor-pointer hover:h-2.5 transition-all group/progress"
        onClick={(e) => {
          const rect = e.currentTarget.getBoundingClientRect();
          const pos = (e.clientX - rect.left) / rect.width;
          if (videoRef.current) videoRef.current.currentTime = pos * duration;
        }}
      >
        {/* Playback Progress Fill */}
        <div 
          className="absolute top-0 left-0 h-full bg-[#F27D26] shadow-[0_0_10px_rgba(242,125,38,0.5)]"
          style={{ width: `${(localTime / (duration || 1)) * 100}%` }}
        />
        
        {/* Detection Progress Fill */}
        {isDetecting && (
          <div 
            className="absolute top-0 left-0 h-full bg-purple-500/30 animate-pulse"
            style={{ width: `${detectionProgress}%` }}
          />
        )}

        {/* Timeline Markers */}
        {markers.map((m, i) => (
          <div 
            key={`${i}-${m.seconds}`} 
            className="absolute top-1/2 w-2.5 h-2.5 -translate-x-1/2 -translate-y-1/2 rounded-full border border-black shadow-xl cursor-pointer hover:scale-150 transition-transform z-10"
            style={{ 
              left: `${(m.seconds / (duration || 1)) * 100}%`,
              backgroundColor: CATEGORY_COLORS[m.category].color 
            }}
            onClick={(e) => {
              e.stopPropagation();
              onSeek(m.seconds);
            }}
            onMouseEnter={() => onSetHoveredMarker(m)}
            onMouseLeave={() => onSetHoveredMarker(null)}
          />
        ))}

        {/* Playhead */}
        <div 
          className="absolute top-1/2 w-3.5 h-3.5 -translate-x-1/2 -translate-y-1/2 bg-white rounded-full shadow-[0_0_10px_rgba(255,255,255,0.5)] pointer-events-none z-20"
          style={{ left: `${(localTime / (duration || 1)) * 100}%` }}
        />
      </div>

      {/* Controls Row */}
      <div className="flex items-center justify-between px-4 py-2 opacity-0 group-hover:opacity-100 transition-opacity duration-300">
        <div className="flex items-center gap-4">
          <button 
            onClick={() => {
              if (videoRef.current) {
                if (videoRef.current.paused) videoRef.current.play();
                else videoRef.current.pause();
              }
            }}
            className="text-white hover:text-[#F27D26] transition-colors"
          >
            {isPaused ? <Play size={18} fill="currentColor" /> : <Square size={16} fill="currentColor" />}
          </button>
          <div className="text-[10px] font-mono text-white/80 select-none">
            <span className="text-white font-bold">{formatTime(localTime)}</span>
            <span className="mx-1 opacity-30">/</span>
            <span className="opacity-50">{formatTime(duration)}</span>
          </div>
        </div>

        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 group/volume">
            <button 
              onClick={() => {
                if (videoRef.current) videoRef.current.muted = !videoRef.current.muted;
              }}
              className="text-white/60 hover:text-white transition-colors"
            >
              {isMuted ? <VolumeX size={16} /> : <Volume2 size={16} />}
            </button>
            <input 
              type="range" 
              min="0" 
              max="1" 
              step="0.1" 
              defaultValue="1"
              onChange={(e) => {
                if (videoRef.current) videoRef.current.volume = parseFloat(e.target.value);
              }}
              className="w-16 h-1 bg-white/10 rounded-full appearance-none cursor-pointer accent-[#F27D26] group-hover/volume:w-20 transition-all"
            />
          </div>
          <button 
            onClick={() => {
              if (document.fullscreenElement) document.exitFullscreen();
              else videoRef.current?.parentElement?.requestFullscreen();
            }}
            className="text-white/60 hover:text-white transition-colors"
          >
            <Maximize2 size={16} />
          </button>
        </div>
      </div>
    </div>
  );
});

const HeaderTime = React.memo(({ 
  videoRef, 
  duration, 
  isRequirementMet, 
  theme 
}: { 
  videoRef: React.RefObject<HTMLVideoElement>, 
  duration: number, 
  isRequirementMet: (type: 'resolution' | 'duration' | 'size') => boolean,
  theme: string
}) => {
  const [localTime, setLocalTime] = useState(0);
  const requestRef = useRef<number>();

  const updateLocalTime = () => {
    if (videoRef.current) {
      setLocalTime(videoRef.current.currentTime);
    }
    requestRef.current = requestAnimationFrame(updateLocalTime);
  };

  useEffect(() => {
    requestRef.current = requestAnimationFrame(updateLocalTime);
    return () => {
      if (requestRef.current) cancelAnimationFrame(requestRef.current);
    };
  }, []);

  const formatTime = (s: number) => {
    return new Date(s * 1000).toISOString().substr(11, 8);
  };

  return (
    <span className={`text-[9px] font-mono ${
      !isRequirementMet('duration')
        ? 'text-red-500 font-bold'
        : 'text-[#F27D26]'
    }`}>
      {formatTime(localTime)} / {formatTime(duration)}
    </span>
  );
});

export default function App() {
  const [comment, setComment] = useState('');
  const [theme, setTheme] = useState<'light' | 'dark'>('dark');
  const [markers, setMarkers] = useState<MarkerData[]>([]);
  const [videoUrl, setVideoUrl] = useState<string | null>(null);
  const [videoList, setVideoList] = useState<VideoFile[]>([]);
  const [selectedVideoId, setSelectedVideoId] = useState<string | null>(null);
  const [duration, setDuration] = useState(0);
  const [playbackRate, setPlaybackRate] = useState(1.0);
  const [lastPlaybackRate, setLastPlaybackRate] = useState(2.0);
  const lastPlaybackRateRef = useRef(2.0);
  const [hoveredMarker, setHoveredMarker] = useState<MarkerData | null>(null);
  const [selectedCategory, setSelectedCategory] = useState<MarkerData['category']>('general');

  const [isAudioImbalanced, setIsAudioImbalanced] = useState(false);
  const [showImbalanceAlert, setShowImbalanceAlert] = useState(false);
  const [hasAddedImbalanceMarker, setHasAddedImbalanceMarker] = useState(false);
  const [isAlertDismissed, setIsAlertDismissed] = useState(false);

  const handleImbalanceChange = (imbalanced: boolean) => {
    setIsAudioImbalanced(imbalanced);
    
    if (imbalanced) {
      // Only show alert if it hasn't been dismissed for this video
      if (!isAlertDismissed) {
        setShowImbalanceAlert(true);
      }
      
      // Only add ONE marker per video
      if (!hasAddedImbalanceMarker && videoRef.current) {
        const videoCurrentTime = videoRef.current.currentTime;
        const timeStr = new Date(videoCurrentTime * 1000).toISOString().substr(11, 8);
        
        // Capture screenshot for the auto-marker
        const captureCanvas = document.createElement('canvas');
        captureCanvas.width = videoRef.current.videoWidth;
        captureCanvas.height = videoRef.current.videoHeight;
        const ctx = captureCanvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(videoRef.current, 0, 0, captureCanvas.width, captureCanvas.height);
        }
        const screenshot = captureCanvas.toDataURL('image/jpeg', 0.6);

        const autoMarker: MarkerData = {
          time: timeStr,
          seconds: videoCurrentTime,
          comment: '系统自动检测：左右声道严重不平衡，请检查音频轨道。',
          category: 'audio',
          color: CATEGORY_COLORS['audio'].prIndex,
          screenshot: screenshot
        };

        setMarkers(prev => [...prev, autoMarker].sort((a, b) => a.seconds - b.seconds));
        setHasAddedImbalanceMarker(true);
        setSuccessMessage('已自动添加音频不平衡批注');
        setTimeout(() => setSuccessMessage(null), 2000);
      }
    }
  };

  const dismissAlert = () => {
    setShowImbalanceAlert(false);
    setIsAlertDismissed(true);
  };

  const processAudio = async (file: File) => {
    setIsAudioLoading(true);
    setAudioBuffer(null);
    try {
      const arrayBuffer = await file.arrayBuffer();
      if (!audioContextRef.current) {
        audioContextRef.current = new (window.AudioContext || (window as any).webkitAudioContext)();
      }
      // Use a fresh context if the old one is closed
      if (audioContextRef.current.state === 'closed') {
        audioContextRef.current = new (window.AudioContext || (window as any).webkitAudioContext)();
      }
      const buffer = await audioContextRef.current.decodeAudioData(arrayBuffer);
      setAudioBuffer(buffer);
    } catch (err) {
      console.error('Audio processing error:', err);
    } finally {
      setIsAudioLoading(false);
    }
  };

  const initAudioNodes = () => {
    if (!videoRef.current) return;
    
    try {
      if (!audioContextRef.current) {
        audioContextRef.current = new (window.AudioContext || (window as any).webkitAudioContext)();
      }

      if (audioContextRef.current.state === 'suspended') {
        audioContextRef.current.resume();
      }

      if (!sourceNodeRef.current) {
        sourceNodeRef.current = audioContextRef.current.createMediaElementSource(videoRef.current);
        
        const splitter = audioContextRef.current.createChannelSplitter(2);
        analyserLRef.current = audioContextRef.current.createAnalyser();
        analyserRRef.current = audioContextRef.current.createAnalyser();
        
        analyserLRef.current.fftSize = 256;
        analyserRRef.current.fftSize = 256;
        
        sourceNodeRef.current.connect(splitter);
        splitter.connect(analyserLRef.current, 0);
        splitter.connect(analyserRRef.current, 1);
        
        sourceNodeRef.current.connect(audioContextRef.current.destination);
        setAudioNodesReady(prev => prev + 1);
      }
    } catch (err) {
      console.error('Failed to initialize audio nodes:', err);
    }
  };

  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  const [quickReplies, setQuickReplies] = useState<Record<string, string[]>>(() => {
    const saved = localStorage.getItem('videoQuickRepliesNested');
    if (saved) return JSON.parse(saved);
    // Migration check for old flat format
    const oldSaved = localStorage.getItem('videoQuickReplies');
    if (oldSaved) {
      const flat = JSON.parse(oldSaved);
      return { ...DEFAULT_QUICK_REPLIES, general: [...new Set([...DEFAULT_QUICK_REPLIES.general, ...flat])] };
    }
    return DEFAULT_QUICK_REPLIES;
  });

  const [activeQuickReplyTab, setActiveQuickReplyTab] = useState<string>('general');

  React.useEffect(() => {
    localStorage.setItem('videoQuickRepliesNested', JSON.stringify(quickReplies));
  }, [quickReplies]);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editingText, setEditingText] = useState('');
  const [showShortcuts, setShowShortcuts] = useState(false);
  const [showDurationCheck, setShowDurationCheck] = useState(false);
  const [verificationRows, setVerificationRows] = useState<{ name: string; expected: string }[]>([{ name: '', expected: '' }]);
  const [verificationResults, setVerificationResults] = useState<{
    name: string;
    expected: string;
    actual: string;
    status: 'match' | 'mismatch' | 'not_found';
  }[]>([]);
  const [verificationHistory, setVerificationHistory] = useState<{
    id: string;
    timestamp: number;
    results: {
      name: string;
      expected: string;
      actual: string;
      status: 'match' | 'mismatch' | 'not_found';
    }[];
  }[]>([]);
  const [showHistory, setShowHistory] = useState(false);
  const [showUnsavedWarning, setShowUnsavedWarning] = useState(false);
  const [pendingUploadEvent, setPendingUploadEvent] = useState<{ event: React.ChangeEvent<HTMLInputElement>, isSingle: boolean } | null>(null);
  const [categoryModal, setCategoryModal] = useState<{ type: 'add' | 'rename' | 'delete', target?: string } | null>(null);
  const [categoryInput, setCategoryInput] = useState('');

  // Check if there are unsaved markers
  const hasUnsavedMarkers = markers.length > 0;

  // Handle beforeunload to warn about unsaved changes
  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (hasUnsavedMarkers) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, [hasUnsavedMarkers]);

  // Load history from localStorage on mount
  useEffect(() => {
    const savedHistory = localStorage.getItem('videoVerificationHistory');
    if (savedHistory) {
      try {
        setVerificationHistory(JSON.parse(savedHistory));
      } catch (e) {
        console.error('Failed to load verification history', e);
      }
    }
  }, []);

  // Save history to localStorage when it changes
  useEffect(() => {
    localStorage.setItem('videoVerificationHistory', JSON.stringify(verificationHistory));
  }, [verificationHistory]);

  const handleVerifyDurations = () => {
    const normalize = (d: string) => {
      if (!d) return '00:00:00';
      const parts = d.replace(/[^\d:]/g, '').split(':').filter(p => p).map(p => p.padStart(2, '0'));
      if (parts.length === 2) return `00:${parts[0]}:${parts[1]}`;
      if (parts.length === 3) return parts.join(':');
      if (parts.length === 1) {
        const secs = parseInt(parts[0]);
        if (isNaN(secs)) return '00:00:00';
        const h = Math.floor(secs / 3600).toString().padStart(2, '0');
        const m = Math.floor((secs % 3600) / 60).toString().padStart(2, '0');
        const s = (secs % 60).toString().padStart(2, '0');
        return `${h}:${m}:${s}`;
      }
      return d;
    };

    const results = verificationRows
      .filter(row => row.name.trim() || row.expected.trim())
      .map(row => {
        const inputName = row.name.trim();
        const inputDuration = row.expected.trim();
        const normalizedExpected = normalize(inputDuration);
        
        const video = videoList.find(v => 
          v.name.toLowerCase().includes(inputName.toLowerCase()) || 
          inputName.toLowerCase().includes(v.name.toLowerCase())
        );
        
        if (!video) {
          return { name: inputName || '未知名称', expected: inputDuration, actual: '--', status: 'not_found' as const };
        }

        const actualDuration = video.duration && video.duration !== '正在加载...' ? video.duration : '00:00:00';
        const normalizedActual = normalize(actualDuration);
        const isMatch = normalizedActual === normalizedExpected;

        return {
          name: video.name,
          expected: inputDuration,
          actual: actualDuration,
          status: isMatch ? 'match' as const : 'mismatch' as const
        };
      });

    setVerificationResults(results);
    
    if (results.length > 0) {
      const newHistoryItem = {
        id: Math.random().toString(36).substr(2, 9),
        timestamp: Date.now(),
        results: results
      };
      setVerificationHistory(prev => [newHistoryItem, ...prev].slice(0, 20));
    }
  };

  const handlePasteTable = (text: string) => {
    const lines = text.split(/\r?\n/).filter(line => line.trim());
    const newRows = lines.map(line => {
      // Split by tab or multiple spaces
      const parts = line.split(/\t|\s{2,}/).map(p => p.trim());
      return {
        name: parts[0] || '',
        expected: parts[1] || ''
      };
    });
    
    setVerificationRows(prev => {
      // If the first row is empty, replace it
      if (prev.length === 1 && !prev[0].name && !prev[0].expected) {
        return [...newRows, { name: '', expected: '' }];
      }
      return [...prev.filter(r => r.name || r.expected), ...newRows, { name: '', expected: '' }];
    });
  };

  const handlePasteColumn = (type: 'name' | 'expected', text: string) => {
    const lines = text.split(/\r?\n/);
    setVerificationRows(prev => {
      const newRows = [...prev];
      lines.forEach((line, i) => {
        if (!newRows[i]) newRows[i] = { name: '', expected: '' };
        newRows[i][type] = line.trim();
      });
      // Add one empty row at the end if needed
      if (newRows.length === lines.length) {
        newRows.push({ name: '', expected: '' });
      }
      return newRows;
    });
  };

  const updateRow = (index: number, field: 'name' | 'expected', value: string) => {
    setVerificationRows(prev => {
      const next = [...prev];
      next[index] = { ...next[index], [field]: value };
      // Add new row if typing in the last row
      if (index === prev.length - 1 && value.trim() !== '') {
        next.push({ name: '', expected: '' });
      }
      return next;
    });
  };

  const removeRow = (index: number) => {
    if (verificationRows.length <= 1) {
      setVerificationRows([{ name: '', expected: '' }]);
      return;
    }
    setVerificationRows(prev => prev.filter((_, i) => i !== index));
  };

  const removeHistoryItem = (id: string) => {
    setVerificationHistory(prev => prev.filter(item => item.id !== id));
  };

  const handleVideoUploadWithWarning = (e: React.ChangeEvent<HTMLInputElement> | { files: FileList | File[], isSingle: boolean, fileHandles?: any[] }, isSingle: boolean) => {
    const files = (e as any).target ? (e as any).target.files : (e as any).files;
    const fileHandles = (e as any).fileHandles;
    const single = (e as any).target ? isSingle : (e as any).isSingle;

    if (hasUnsavedMarkers) {
      setPendingUploadEvent({ files, isSingle: single, fileHandles });
      setShowUnsavedWarning(true);
    } else {
      handleVideoUploadManual(files, single, fileHandles);
    }
  };

  const confirmDiscardAndUpload = () => {
    if (pendingUploadEvent) {
      const { files, isSingle, fileHandles } = pendingUploadEvent;
      // Manually trigger handleVideoUpload with the stored files
      handleVideoUploadManual(files, isSingle, fileHandles);
      setPendingUploadEvent(null);
    }
    setShowUnsavedWarning(false);
  };

  const handleVideoUploadManual = async (files: FileList | File[] | null, replace: boolean = false, fileHandles?: any[]) => {
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
            duration: '正在加载...',
            markers: [],
            fileHandle: fileHandles ? fileHandles[i] : undefined
          });
        }
      }
      
      if (newVideos.length > 0) {
        let updatedList: VideoFile[];
        if (replace) {
          updatedList = newVideos;
          setVideoList(newVideos);
          const firstVideo = newVideos[0];
          setMarkers([]);
          setVideoUrl(firstVideo.url);
          setSelectedVideoId(firstVideo.id);
          setVideoMetadata({
            name: firstVideo.name,
            size: firstVideo.size,
            resolution: '正在加载...',
            duration: '正在加载...',
            fileHandle: firstVideo.fileHandle
          });
          processAudio(firstVideo.file);
        } else {
          updatedList = [...videoList, ...newVideos];
          setVideoList(prev => [...prev, ...newVideos]);
          if (!selectedVideoId) {
            const firstVideo = newVideos[0];
            selectVideo(firstVideo.id);
          }
        }

        // Load metadata for all new videos
        newVideos.forEach(video => {
          const tempVideo = document.createElement('video');
          tempVideo.src = video.url;
          tempVideo.onloadedmetadata = () => {
            const dur = new Date(tempVideo.duration * 1000).toISOString().substr(11, 8);
            const res = `${tempVideo.videoWidth}x${tempVideo.videoHeight}`;
            
            setVideoList(prev => prev.map(v => 
              v.id === video.id ? { ...v, duration: dur, resolution: res } : v
            ));

            if (selectedVideoId === video.id || (replace && video.id === newVideos[0].id)) {
              setVideoMetadata(prev => prev ? { ...prev, duration: dur, resolution: res } : null);
            }
          };
        });
      }
    }
  };

  const videoInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);
  const replaceVideoInputRef = useRef<HTMLInputElement>(null);

  const triggerVideoUpload = (replace: boolean = false) => {
    if (replace) {
      replaceVideoInputRef.current?.click();
    } else {
      videoInputRef.current?.click();
    }
  };

  const triggerFolderUpload = () => {
    folderInputRef.current?.click();
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    
    const items = e.dataTransfer.items;
    if (items) {
      const files: File[] = [];
      const fileHandles: any[] = []; // Handles are generally not available through standard drop
      
      const traverseEntry = async (entry: any) => {
        if (entry.isFile) {
          const file = await new Promise<File>((resolve) => entry.file(resolve));
          if (file.type.startsWith('video/')) {
            files.push(file);
          }
        } else if (entry.isDirectory) {
          const reader = entry.createReader();
          const readAllEntries = async (dirReader: any) => {
            const results: any[] = [];
            let readBatch = async () => {
              const entries = await new Promise<any[]>((resolve) => dirReader.readEntries(resolve));
              if (entries.length > 0) {
                results.push(...entries);
                await readBatch();
              }
            };
            await readBatch();
            return results;
          };
          const entries = await readAllEntries(reader);
          for (const subEntry of entries) {
            await traverseEntry(subEntry);
          }
        }
      };

      const scanPromises = [];
      for (let i = 0; i < items.length; i++) {
        const entry = items[i].webkitGetAsEntry();
        if (entry) {
          scanPromises.push(traverseEntry(entry));
        }
      }
      
      await Promise.all(scanPromises);
      
      if (files.length > 0) {
        handleVideoUploadWithWarning({ files, isSingle: false }, false);
      }
    }
  };

  const [audioBuffer, setAudioBuffer] = useState<AudioBuffer | null>(null);
  const [isAudioLoading, setIsAudioLoading] = useState(false);
  const audioContextRef = React.useRef<AudioContext | null>(null);
  const analyserLRef = React.useRef<AnalyserNode | null>(null);
  const analyserRRef = React.useRef<AnalyserNode | null>(null);
  const sourceNodeRef = React.useRef<MediaElementAudioSourceNode | null>(null);

  const [audioNodesReady, setAudioNodesReady] = useState(0);

  const [isDetecting, setIsDetecting] = useState(false);
  const [detectionProgress, setDetectionProgress] = useState(0);
  const stopDetectionRef = React.useRef(false);
  const [drawingMode, setDrawingMode] = useState<'none' | 'brush' | 'arrow' | 'rect'>('none');
  const [drawingColor, setDrawingColor] = useState('#F27D26');
  const [isHeaderHovered, setIsHeaderHovered] = useState(false);
  const [isDrawing, setIsDrawing] = useState(false);
  const [startPos, setStartPos] = useState({ x: 0, y: 0 });
  const handleAddReviewRef = React.useRef<() => void>(() => {});
  
  const [isDragging, setIsDragging] = useState(false);
  
  const [videoMetadata, setVideoMetadata] = useState<{
    name: string;
    size: string;
    resolution: string;
    duration?: string;
    fileHandle?: any;
  } | null>(null);
  
  const [showRequirementsModal, setShowRequirementsModal] = useState(false);
  const [isResolutionDropdownOpen, setIsResolutionDropdownOpen] = useState(false);
  const [requirements, setRequirements] = useState<{
    resolution: string;
    duration: number | null;
    size: number | null;
  }>({
    resolution: '',
    duration: null,
    size: null,
  });

  const isRequirementMet = (type: 'resolution' | 'duration' | 'size') => {
    if (!videoMetadata || !videoUrl) return true;
    
    switch (type) {
      case 'resolution':
        if (!requirements.resolution) return true;
        return videoMetadata.resolution === requirements.resolution;
      case 'duration':
        if (requirements.duration === null) return true;
        return duration <= requirements.duration * 60;
      case 'size':
        if (requirements.size === null) return true;
        const sizeNum = parseFloat(videoMetadata.size.split(' ')[0]);
        return sizeNum <= requirements.size;
      default:
        return true;
    }
  };

  const videoRef = React.useRef<HTMLVideoElement>(null);
  const drawingCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const tempCanvasRef = React.useRef<HTMLCanvasElement>(null);
  const commentInputRef = React.useRef<HTMLTextAreaElement>(null);
  const lastAddReviewTimeRef = React.useRef<number>(0);
  // Handle Video Upload
  const handleVideoUpload = async (e: React.ChangeEvent<HTMLInputElement>, replace: boolean = false) => {
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
            duration: '正在加载...',
            markers: []
          });
        }
      }
      
      if (newVideos.length > 0) {
        let updatedList: VideoFile[];
        if (replace) {
          updatedList = newVideos;
          setVideoList(newVideos);
          const firstVideo = newVideos[0];
          setMarkers([]);
          setVideoUrl(firstVideo.url);
          setSelectedVideoId(firstVideo.id);
          setVideoMetadata({
            name: firstVideo.name,
            size: firstVideo.size,
            resolution: '正在加载...',
            duration: '正在加载...'
          });
          processAudio(firstVideo.file);
        } else {
          updatedList = [...videoList, ...newVideos];
          setVideoList(prev => [...prev, ...newVideos]);
          if (!selectedVideoId) {
            const firstVideo = newVideos[0];
            selectVideo(firstVideo.id);
            processAudio(firstVideo.file);
          }
        }

        // Automatically load metadata for all new videos
        newVideos.forEach(async (v) => {
          try {
            const meta = await new Promise<{ duration: string, resolution: string }>((resolve) => {
              const video = document.createElement('video');
              video.preload = 'metadata';
              video.onloadedmetadata = () => {
                const duration = new Date(video.duration * 1000).toISOString().substr(11, 8);
                const resolution = `${video.videoWidth}x${video.videoHeight}`;
                URL.revokeObjectURL(video.src);
                resolve({ duration, resolution });
              };
              video.onerror = () => resolve({ duration: '未知', resolution: '未知' });
              video.src = URL.createObjectURL(v.file);
            });

            setVideoList(prev => prev.map(item => 
              item.id === v.id ? { ...item, ...meta } : item
            ));

            // Also update current metadata if this is the selected video
            if (v.id === selectedVideoId) {
              setVideoMetadata(prev => prev ? { ...prev, ...meta } : null);
            }
          } catch (err) {
            console.error('Metadata loading failed:', err);
          }
        });

        setSuccessMessage(replace ? '已替换当前视频' : `成功导入 ${newVideos.length} 个视频`);
        setTimeout(() => setSuccessMessage(null), 3000);
      }
    }
    // Reset input value to allow re-uploading the same file
    e.target.value = '';
  };

  // Sync Video Time
  const handleTimeUpdate = () => {
    // Current time is handled locally in VideoControls child component for performance
  };

  const handlePlay = () => {
    initAudioNodes();
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
    initAudioNodes();
    if (videoRef.current && drawingCanvasRef.current && tempCanvasRef.current) {
      const video = videoRef.current;
      setDuration(video.duration);
      
      const resolution = `${video.videoWidth} x ${video.videoHeight}`;
      const durationStr = new Date(video.duration * 1000).toISOString().substr(11, 8);
      
      // Update resolution and duration in metadata
      if (videoMetadata) {
        setVideoMetadata({
          ...videoMetadata,
          resolution: resolution,
          duration: durationStr
        });
      }

      // Update in videoList
      if (selectedVideoId) {
        setVideoList(prev => prev.map(v => 
          v.id === selectedVideoId ? { ...v, resolution, duration: durationStr } : v
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

  const seekTo = React.useCallback((seconds: number) => {
    if (videoRef.current) {
      videoRef.current.currentTime = seconds;
      videoRef.current.pause(); // Pause at the point of review
    }
  }, []);

  const deleteMarker = React.useCallback((index: number) => {
    setMarkers(prev => {
      const next = [...prev];
      next.splice(index, 1);
      return next;
    });
  }, []);

  const selectVideo = React.useCallback((id: string) => {
    // Save current markers to the list before switching
    if (selectedVideoId) {
      setVideoList(prev => prev.map(v => 
        v.id === selectedVideoId ? { ...v, markers: [...markers] } : v
      ));
    }

    const video = videoList.find(v => v.id === id);
    if (video) {
      setSelectedVideoId(id);
      setVideoUrl(video.url);
      setMarkers(video.markers || []);
      setVideoMetadata({
        name: video.name,
        size: video.size,
        resolution: video.resolution,
        duration: video.duration,
        fileHandle: video.fileHandle
      });
      // Process audio for visualization
      processAudio(video.file);
      setHasAddedImbalanceMarker(false);
      setIsAlertDismissed(false);
    }
  }, [videoList, selectedVideoId, markers]);

  const handleDeleteVideo = React.useCallback((id: string) => {
    if (window.confirm('确定要从列表中移除此视频吗？')) {
      setVideoList(prev => prev.filter(v => v.id !== id));
      if (selectedVideoId === id) {
        setVideoUrl(null);
        setSelectedVideoId(null);
        setMarkers([]);
        setVideoMetadata(null);
      }
    }
  }, [selectedVideoId]);

  const removeQuickReply = React.useCallback((tab: string, index: number) => {
    setQuickReplies(prev => {
      const current = [...(prev[tab] || [])];
      current.splice(index, 1);
      return { ...prev, [tab]: current };
    });
  }, []);

  const handleAddQuickReplyToComment = React.useCallback((reply: string) => {
    // Try to auto-map category for color matching if category name matches CATEGORY_COLORS labels
    const matchedCategory = Object.keys(CATEGORY_COLORS).find(k => 
      CATEGORY_COLORS[k as keyof typeof CATEGORY_COLORS].label === activeQuickReplyTab
    );
    if (matchedCategory) {
       setSelectedCategory(matchedCategory as keyof typeof CATEGORY_COLORS);
    }
    setTimeout(() => handleAddReview(reply), 0);
  }, [activeQuickReplyTab]);

  const handleImportQuickReplies = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.txt,.json,.xlsx,.xls';
    input.onchange = async (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (file) {
        try {
          if (file.name.endsWith('.json')) {
            const text = await file.text();
            const data = JSON.parse(text);
            setQuickReplies(prev => ({ ...prev, ...data }));
            setSuccessMessage('导入配置成功');
          } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];
            
            const newReplies: Record<string, string[]> = {};
            jsonData.forEach(row => {
              const cat = row.Category || row.分类 || '未分类';
              const content = row.Content || row.话术 || row.内容;
              if (content) {
                if (!newReplies[cat]) newReplies[cat] = [];
                newReplies[cat].push(String(content));
              }
            });

            if (Object.keys(newReplies).length > 0) {
              setQuickReplies(prev => ({ ...prev, ...newReplies }));
              setSuccessMessage(`已从Excel导入 ${Object.keys(newReplies).length} 个分类`);
            }
          } else {
            const text = await file.text();
            const lines = text.split(/\r?\n/).filter(line => line.trim());
            if (lines.length > 0) {
              setQuickReplies(prev => ({
                ...prev,
                [activeQuickReplyTab]: [...new Set([...(prev[activeQuickReplyTab] || []), ...lines])]
              }));
              setSuccessMessage(`已导入 ${lines.length} 条快捷回复到当前分类`);
            }
          }
          setTimeout(() => setSuccessMessage(null), 2000);
        } catch (err) {
          console.error('Failed to import quick replies', err);
          setSuccessMessage('导入失败，请检查文件格式');
          setTimeout(() => setSuccessMessage(null), 2000);
        }
      }
    };
    input.click();
  };

  const handleExportQuickReplies = () => {
    // Export to Excel
    const data: any[] = [];
    Object.entries(quickReplies).forEach(([cat, replies]) => {
      (replies as string[]).forEach(reply => {
        data.push({ 分类: cat, 话术: reply });
      });
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '快捷话术');
    
    // Auto-size columns
    const maxCategoryWidth = Math.max(...Object.keys(quickReplies).map(k => k.length), 10);
    worksheet['!cols'] = [{ wch: maxCategoryWidth + 5 }, { wch: 50 }];

    XLSX.writeFile(workbook, `快捷话术_${new Date().toISOString().split('T')[0]}.xlsx`);
    
    setSuccessMessage('配置已以Excel形式导出');
    setTimeout(() => setSuccessMessage(null), 2000);
  };

  const addCategory = () => {
    setCategoryInput('');
    setCategoryModal({ type: 'add' });
  };

  const confirmAddCategory = () => {
    const name = categoryInput.trim();
    if (name) {
      if (quickReplies[name]) {
        alert('分类已存在');
        return;
      }
      setQuickReplies(prev => ({ ...prev, [name]: [] }));
      setActiveQuickReplyTab(name);
      setCategoryModal(null);
    }
  };

  const deleteCategory = (cat: string) => {
    setCategoryModal({ type: 'delete', target: cat });
  };

  const confirmDeleteCategory = () => {
    if (categoryModal?.target) {
      const cat = categoryModal.target;
      const newReplies = { ...quickReplies };
      delete newReplies[cat];
      setQuickReplies(newReplies);
      if (activeQuickReplyTab === cat) {
        setActiveQuickReplyTab(Object.keys(newReplies)[0] || '');
      }
      setCategoryModal(null);
    }
  };

  const renameCategory = (oldName: string) => {
    setCategoryInput(oldName);
    setCategoryModal({ type: 'rename', target: oldName });
  };

  const confirmRenameCategory = () => {
    if (categoryModal?.target) {
      const oldName = categoryModal.target;
      const newName = categoryInput.trim();
      if (newName && newName !== oldName) {
        if (quickReplies[newName]) {
          alert('新分类名称已存在');
          return;
        }
        const newReplies = { ...quickReplies };
        newReplies[newName] = newReplies[oldName];
        delete newReplies[oldName];
        setQuickReplies(newReplies);
        if (activeQuickReplyTab === oldName) {
          setActiveQuickReplyTab(newName);
        }
        setCategoryModal(null);
      }
    }
  };

  // Add Review Comment
  const handleAddReview = (overrideComment?: string) => {
    if (!videoRef.current) return;

    // Throttle: 1 second between comments
    const now = Date.now();
    if (now - lastAddReviewTimeRef.current < 1000) {
      return;
    }

    const activeComment = typeof overrideComment === 'string' ? overrideComment : comment;

    // Check if both comment and canvas are empty
    const isCommentEmpty = !activeComment.trim();
    let isCanvasEmpty = true;
    if (drawingCanvasRef.current) {
      const ctx = drawingCanvasRef.current.getContext('2d');
      if (ctx) {
        const imageData = ctx.getImageData(0, 0, drawingCanvasRef.current.width, drawingCanvasRef.current.height);
        const data = imageData.data;
        for (let i = 0; i < data.length; i += 4) {
          if (data[i + 3] > 0) { // Check alpha channel
            isCanvasEmpty = false;
            break;
          }
        }
      }
    }

    if (isCommentEmpty && isCanvasEmpty) {
      setSuccessMessage('请输入批注内容或在画面上进行标注');
      setTimeout(() => setSuccessMessage(null), 2000);
      return;
    }

    const videoCurrentTime = videoRef.current.currentTime;

    // Check for duplicates (same time and same comment) - use small epsilon for float comparison
    const isDuplicate = markers.some(m => 
      Math.abs(m.seconds - videoCurrentTime) < 0.1 && 
      m.comment.trim() === activeComment.trim() &&
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
    const timeStr = new Date(videoCurrentTime * 1000).toISOString().substr(11, 8);

    const newMarker: MarkerData = {
      time: timeStr,
      seconds: videoCurrentTime,
      comment: activeComment,
      category: selectedCategory,
      color: CATEGORY_COLORS[selectedCategory].prIndex,
      screenshot: screenshot
    };

    setMarkers(prev => [...prev, newMarker].sort((a, b) => a.seconds - b.seconds));
    if (typeof overrideComment !== 'string') setComment('');
    clearCanvas();
    
    // Blur the input to allow keyboard shortcuts (like space for play/pause) to work
    if (commentInputRef.current) {
      commentInputRef.current.blur();
    }
    
    lastAddReviewTimeRef.current = Date.now();
    setSuccessMessage('批注已添加');
    setTimeout(() => setSuccessMessage(null), 2000);
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

  // Keyboard Shortcuts (Standard NLE Style)
  React.useEffect(() => {
    handleAddReviewRef.current = handleAddReview;
    lastPlaybackRateRef.current = lastPlaybackRate;
  });

  React.useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Global Ctrl+Enter shortcut
      if (e.ctrlKey && e.key === 'Enter') {
        e.preventDefault();
        handleAddReviewRef.current();
        return;
      }

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
          e.stopPropagation();
          // If something is focused (like a button), blur it to prevent double-triggering
          if (document.activeElement instanceof HTMLElement && 
              !(document.activeElement instanceof HTMLInputElement) && 
              !(document.activeElement instanceof HTMLTextAreaElement)) {
            document.activeElement.blur();
          }
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
        case 'KeyD':
          e.preventDefault();
          video.pause();
          video.currentTime = Math.max(0, video.currentTime - frameTime);
          break;
        case 'Period': // Next Frame
        case 'KeyF':
          e.preventDefault();
          video.pause();
          video.currentTime = Math.min(video.duration, video.currentTime + frameTime);
          break;
        case 'Enter':
        case 'NumpadEnter':
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
          const newRateC = Math.min(3.0, Math.round((video.playbackRate + 0.1) * 10) / 10);
          video.playbackRate = newRateC;
          setPlaybackRate(newRateC);
          if (newRateC !== 1.0) setLastPlaybackRate(newRateC);
          setSuccessMessage(`播放速度: ${newRateC.toFixed(1)}x`);
          setTimeout(() => setSuccessMessage(null), 1000);
          break;
        case 'KeyX': // Slow Down
          e.preventDefault();
          const newRateX = Math.max(0.1, Math.round((video.playbackRate - 0.1) * 10) / 10);
          video.playbackRate = newRateX;
          setPlaybackRate(newRateX);
          if (newRateX !== 1.0) setLastPlaybackRate(newRateX);
          setSuccessMessage(`播放速度: ${newRateX.toFixed(1)}x`);
          setTimeout(() => setSuccessMessage(null), 1000);
          break;
        case 'KeyZ': // Toggle Speed
          e.preventDefault();
          if (video.playbackRate !== 1.0) {
            setLastPlaybackRate(video.playbackRate);
            video.playbackRate = 1.0;
            setPlaybackRate(1.0);
            setSuccessMessage('播放速度: 1.0x (重置)');
          } else {
            const targetRate = lastPlaybackRateRef.current;
            video.playbackRate = targetRate;
            setPlaybackRate(targetRate);
            setSuccessMessage(`播放速度: ${targetRate.toFixed(1)}x (切换)`);
          }
          setTimeout(() => setSuccessMessage(null), 1000);
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

  // Black Frame Detection Logic (High Performance with WebCodecs)
  const detectBlackFrames = async () => {
    if (!videoUrl || !selectedVideoId) return;
    const currentVideo = videoList.find(v => v.id === selectedVideoId);
    if (!currentVideo || !currentVideo.file) return;
    
    if (isDetecting) {
      stopDetectionRef.current = true;
      return;
    }

    setIsDetecting(true);
    stopDetectionRef.current = false;
    setDetectionProgress(0);
    setSuccessMessage('正在初始化并行检测引擎...');

    try {
      if (!window.VideoDecoder) throw new Error('当前浏览器不支持 WebCodecs');

      const file = currentVideo.file;
      const isMP4 = file.name.toLowerCase().endsWith('.mp4') || file.type === 'video/mp4';
      if (!isMP4) throw new Error('非 MP4 容器，无法使用并行模式');

      const mp4boxFile = MP4Box.createFile();
      let videoTrack: any = null;
      let isReady = false;
      const threshold = 16;
      const numWorkers = 4;
      const decoders: VideoDecoder[] = [];
      const canvases: HTMLCanvasElement[] = [];
      const contexts: CanvasRenderingContext2D[] = [];
      const currentSequences: ({ start: number, count: number, screenshot: string } | null)[] = new Array(numWorkers).fill(null);
      const detectedMarkers: MarkerData[] = [];
      let samplesProcessed = 0;

      await new Promise<void>((resolve, reject) => {
        const timeout = setTimeout(() => {
          if (!isReady) reject(new Error('视频解析超时'));
        }, 20000);

        mp4boxFile.onReady = (info: any) => {
          isReady = true;
          clearTimeout(timeout);
          videoTrack = info.videoTracks[0];
          if (!videoTrack) return reject(new Error('未找到视频轨道'));

          setSuccessMessage(`[并行模式] 正在启动 4 线程检测...`);

          // Initialize workers
          const track = mp4boxFile.getTrackById(videoTrack.id);
          const entry = track.mdia.minf.stbl.stsd.entries[0] as any;
          const box = entry.avcC || entry.hvcC || entry.vpcC;
          let description: Uint8Array | undefined;
          if (box) {
            const stream = new (MP4Box as any).DataStream(undefined, 0, (MP4Box as any).DataStream.BIG_ENDIAN);
            box.write(stream);
            description = new Uint8Array(stream.buffer, 8);
          }

          const config: VideoDecoderConfig = {
            codec: videoTrack.codec,
            codedWidth: videoTrack.track_width,
            codedHeight: videoTrack.track_height,
            description
          };

          for (let i = 0; i < numWorkers; i++) {
            const canvas = document.createElement('canvas');
            canvas.width = 16;
            canvas.height = 9;
            const ctx = canvas.getContext('2d', { willReadFrequently: true })!;
            canvases.push(canvas);
            contexts.push(ctx);

            const decoder = new VideoDecoder({
              output: (frame) => {
                const workerIdx = i;
                const time = frame.timestamp / 1000000;
                ctx.drawImage(frame, 0, 0, 16, 9);
                const data = ctx.getImageData(0, 0, 16, 9).data;
                let brightness = 0;
                for (let j = 0; j < data.length; j += 4) brightness += (data[j] + data[j+1] + data[j+2]) / 3;
                const avg = brightness / (data.length / 4);

                if (avg < threshold) {
                  if (!currentSequences[workerIdx]) {
                    currentSequences[workerIdx] = { start: time, count: 1, screenshot: canvas.toDataURL('image/jpeg', 0.5) };
                  } else {
                    currentSequences[workerIdx]!.count++;
                  }
                } else if (currentSequences[workerIdx]) {
                  const seq = currentSequences[workerIdx]!;
                  detectedMarkers.push({
                    time: new Date(seq.start * 1000).toISOString().substr(11, 8),
                    seconds: seq.start,
                    comment: `检测到黑帧序列 (持续 ${seq.count} 帧)`,
                    category: 'technical',
                    color: CATEGORY_COLORS['technical'].prIndex,
                    screenshot: seq.screenshot,
                    isConfirmed: false
                  });
                  currentSequences[workerIdx] = null;
                }
                frame.close();
              },
              error: (e) => console.error(`Worker ${i} error:`, e)
            });
            decoder.configure(config);
            decoders.push(decoder);
          }

          (mp4boxFile as any).setExtractionConfig(videoTrack.id, null, { nb_samples: 500 });
          mp4boxFile.start();
        };

        mp4boxFile.onSamples = (id, user, samples) => {
          if (!videoTrack || !videoTrack.nb_samples) return;
          
          for (const sample of samples) {
            if (stopDetectionRef.current) break;
            
            // Distribute samples by segments to preserve sequences
            const total = videoTrack.nb_samples;
            const workerIdx = Math.min(numWorkers - 1, Math.floor((samplesProcessed / total) * numWorkers));
            
            decoders[workerIdx].decode(new EncodedVideoChunk({
              type: sample.is_sync ? 'key' : 'delta',
              timestamp: (sample.cts * 1000000) / sample.timescale,
              duration: (sample.duration * 1000000) / sample.timescale,
              data: sample.data
            }));
            
            samplesProcessed++;
          }
          const progress = Math.round((samplesProcessed / videoTrack.nb_samples) * 100);
          setDetectionProgress(Math.min(progress, 99));
          setSuccessMessage(`正在并行解码检测 (${progress}%)...`);
        };

        const reader = file.stream().getReader();
        let offset = 0;
        const read = async () => {
          try {
            while (true) {
              if (stopDetectionRef.current) break;

              // Backpressure: check if any decoder is overwhelmed
              const isOverwhelmed = decoders.some(d => d.decodeQueueSize > 100);
              if (isOverwhelmed) {
                await new Promise(r => setTimeout(r, 50));
                continue;
              }

              const { done, value } = await reader.read();
              if (done) {
                mp4boxFile.flush();
                // Wait for all decoders to finish
                await Promise.all(decoders.map(d => d.flush()));
                
                // Handle remaining sequences
                currentSequences.forEach((seq) => {
                  if (seq) {
                    detectedMarkers.push({
                      time: new Date(seq.start * 1000).toISOString().substr(11, 8),
                      seconds: seq.start,
                      comment: `检测到黑帧序列 (持续 ${seq.count} 帧)`,
                      category: 'technical',
                      color: CATEGORY_COLORS['technical'].prIndex,
                      screenshot: seq.screenshot,
                      isConfirmed: false
                    });
                  }
                });

                setDetectionProgress(100);
                resolve();
                break;
              }

              const buffer = value.buffer as ArrayBuffer;
              (buffer as any).fileStart = offset;
              mp4boxFile.appendBuffer(buffer as any);
              offset += buffer.byteLength;

              if (!isReady) {
                const readProgress = Math.round((offset / file.size) * 100);
                setSuccessMessage(`正在读取视频索引 (${readProgress}%, ${Math.round(offset / 1024 / 1024)}MB)...`);
                await new Promise(r => setTimeout(r, 0));
              }
            }
          } catch (e) { reject(e); }
        };
        read();
      });

      const finalMarkers = [];
      if (detectedMarkers.length > 0) {
        const sorted = detectedMarkers.sort((a, b) => a.seconds - b.seconds);
        let currentGroup = { ...sorted[0] };
        let lastSeconds = currentGroup.seconds;
        
        for (let i = 1; i < sorted.length; i++) {
          // Sliding window: if gap between THIS and PREVIOUS is <= 15s, merge
          if (sorted[i].seconds - lastSeconds <= 15) {
            lastSeconds = sorted[i].seconds;
            currentGroup.comment = `检测到黑帧异常区域 (已自动汇总相邻项)`;
          } else {
            finalMarkers.push(currentGroup);
            currentGroup = { ...sorted[i] };
            lastSeconds = currentGroup.seconds;
          }
        }
        finalMarkers.push(currentGroup);
      }

      setMarkers(prev => [...prev, ...finalMarkers].sort((a, b) => a.seconds - b.seconds));
      setSuccessMessage(`并行检测完成！共发现 ${finalMarkers.length} 处异常区域。`);
      setTimeout(() => setSuccessMessage(null), 3000);
    } catch (error) {
      console.error('Parallel detection failed:', error);
      setSuccessMessage('并行模式启动失败，正在切换至兼容模式...');
      setTimeout(() => detectBlackFramesCompatible(), 1000);
    } finally {
      setIsDetecting(false);
    }
  };

  // Legacy Black Frame Detection Logic (Fallback)
  const detectBlackFramesCompatible = async () => {
    if (!videoUrl || !videoRef.current) return;
    
    setIsDetecting(true);
    setDetectionProgress(0);
    setSuccessMessage('正在使用兼容模式检测...');

    const video = document.createElement('video');
    video.src = videoUrl;
    video.crossOrigin = 'anonymous';
    video.muted = true;
    
    try {
      await new Promise((resolve, reject) => {
        video.onloadedmetadata = resolve;
        video.onerror = () => reject(new Error('视频加载失败'));
        setTimeout(() => reject(new Error('加载超时')), 10000);
      });

      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      if (!ctx) throw new Error('无法创建画布上下文');

      const duration = video.duration;
      const interval = 0.5; 
      const threshold = 12; 
      const detectedMarkers: MarkerData[] = [];
      
      canvas.width = 160; 
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
              comment: '检测到疑似黑帧 (兼容模式)',
              category: 'technical',
              color: CATEGORY_COLORS['technical'].prIndex,
              screenshot: screenshot,
              isConfirmed: false
            });
          }
        }
        
        setDetectionProgress(Math.round((time / duration) * 100));
      }

      const finalMarkers = [];
      if (detectedMarkers.length > 0) {
        const sorted = detectedMarkers.sort((a, b) => a.seconds - b.seconds);
        let currentGroup = { ...sorted[0] };
        let lastSeconds = currentGroup.seconds;
        
        for (let i = 1; i < sorted.length; i++) {
          // Sliding window: if gap between THIS and PREVIOUS is <= 15s, merge
          if (sorted[i].seconds - lastSeconds <= 15) {
            lastSeconds = sorted[i].seconds;
            currentGroup.comment = `检测到黑帧异常区域 (已自动汇总相邻项)`;
          } else {
            finalMarkers.push(currentGroup);
            currentGroup = { ...sorted[i] };
            lastSeconds = currentGroup.seconds;
          }
        }
        finalMarkers.push(currentGroup);
      }

      if (finalMarkers.length > 0) {
        setMarkers(prev => [...prev, ...finalMarkers].sort((a, b) => a.seconds - b.seconds));
        setSuccessMessage(`检测完成，共发现 ${finalMarkers.length} 处异常区域`);
      } else {
        setSuccessMessage('检测完成，未发现黑帧');
      }
    } catch (err) {
      console.error('Legacy Detection error:', err);
      setSuccessMessage('检测出错');
    } finally {
      setIsDetecting(false);
      setDetectionProgress(0);
      stopDetectionRef.current = false;
      setTimeout(() => setSuccessMessage(null), 3000);
    }
  };

  // Export Logic
  const saveFile = async (blob: Blob, suggestedName: string, fileHandle?: any) => {
    if ('showSaveFilePicker' in window) {
      try {
        const handle = await (window as any).showSaveFilePicker({
          suggestedName,
          startIn: fileHandle,
          types: [{
            description: suggestedName.endsWith('.docx') ? 'Word Document' : 'Excel Spreadsheet',
            accept: suggestedName.endsWith('.docx') 
              ? {'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx']}
              : {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']},
          }],
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return;
      } catch (err: any) {
        if (err.name === 'AbortError') return;
        console.error('File System Access API failed:', err);
      }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = suggestedName;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportReviewCommentsToExcel = async () => {
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
    const fileName = videoMetadata 
      ? `审核报告_${videoMetadata.name.split('.')[0]}_${new Date().toISOString().split('T')[0]}.xlsx`
      : `视频审核报告_${new Date().toLocaleDateString()}.xlsx`;
    
    await saveFile(blob, fileName, videoMetadata?.fileHandle);
  };

  const exportReviewCommentsToWord = async () => {
    if (markers.length === 0) return;

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

    for (const m of markers) {
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
      sections: [
        {
          children: [
            new Paragraph({
              text: videoMetadata ? `视频审核报告 - ${videoMetadata.name}` : "视频审核报告",
              heading: HeadingLevel.HEADING_1,
              spacing: { after: 200 },
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: tableRows,
            }),
          ],
        },
      ],
    });

    const buffer = await Packer.toBlob(doc);
    const fileName = videoMetadata 
      ? `审核报告_${videoMetadata.name.split('.')[0]}_${new Date().toISOString().split('T')[0]}.docx`
      : `视频审核报告_${new Date().toLocaleDateString()}.docx`;
    
    await saveFile(buffer, fileName, videoMetadata?.fileHandle);
  };

  const exportBatchReport = async () => {
    if (videoList.length === 0) return;
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('视频列表汇总');
    
    worksheet.columns = [
      { header: '视频名称', key: 'name', width: 40 },
      { header: '视频时长', key: 'duration', width: 15 },
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
        duration: v.duration || '未知',
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
    const fileName = `批量审核汇总_${new Date().toLocaleDateString()}.xlsx`;
    
    await saveFile(blob, fileName, videoMetadata?.fileHandle);
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
    const fileName = `全量视频批注汇总_${new Date().toLocaleDateString()}.xlsx`;
    
    await saveFile(blob, fileName, videoMetadata?.fileHandle);
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
    const fileName = `全量视频批注汇总_${new Date().toLocaleDateString()}.docx`;
    
    await saveFile(buffer, fileName, videoMetadata?.fileHandle);
  };

  return (
    <div className={`min-h-screen transition-colors duration-300 ${theme === 'dark' ? 'bg-[#0A0A0A] text-[#E0E0E0]' : 'bg-[#F8F9FA] text-[#1A1A1A]'} font-sans selection:bg-[#F27D26] overflow-hidden`}>
      <div className="max-w-[1600px] mx-auto p-4 md:p-6 h-screen flex flex-col relative">
        
        {/* Header Auto-hide Container */}
        <div 
          onMouseEnter={() => setIsHeaderHovered(true)}
          onMouseLeave={() => setIsHeaderHovered(false)}
          className="absolute top-0 left-0 right-0 z-[100] h-16 group"
        >
          {/* Invisible trigger area that's always at the top */}
          <div className="absolute top-0 left-0 right-0 h-2" />
          
          <motion.header 
            initial={false}
            animate={{ 
              y: isHeaderHovered ? 0 : -80,
              opacity: isHeaderHovered ? 1 : 0
            }}
            transition={{ type: 'spring', damping: 25, stiffness: 120 }}
            className={`flex items-center justify-between px-8 py-4 rounded-b-3xl border-x border-b shadow-2xl backdrop-blur-xl ${theme === 'dark' ? 'bg-black/60 border-white/5' : 'bg-white/90 border-gray-200/50 shadow-sm'}`}
          >
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-[#F27D26] rounded-xl flex items-center justify-center shadow-lg shadow-[#F27D26]/20">
                <Play className="text-white fill-current" size={20} />
              </div>
              <div>
                <h1 className={`text-xl font-bold tracking-tight ${theme === 'dark' ? 'text-white' : 'text-gray-900'}`}>ReviewFlow <span className="text-[#F27D26] text-xs font-normal ml-2">Studio v1.0</span></h1>
                <p className={`text-[9px] uppercase tracking-[0.2em] font-mono ${theme === 'dark' ? 'opacity-30' : 'text-gray-400'}`}>Professional Video Review Platform</p>
              </div>
            </div>
            
            <div className="flex items-center gap-4">
              <button 
                onClick={() => setShowDurationCheck(true)}
                className={`p-2 rounded-lg transition-all ${theme === 'dark' ? 'hover:bg-white/5 text-white/40' : 'hover:bg-gray-100 text-gray-400'}`}
                title="时长核对"
              >
                <FileStack size={18} />
              </button>
              
              <button 
                onClick={() => setShowShortcuts(true)}
                className={`p-2 rounded-lg transition-all ${theme === 'dark' ? 'hover:bg-white/5 text-white/40' : 'hover:bg-gray-100 text-gray-400'}`}
                title="键盘快捷键"
              >
                <Keyboard size={18} />
              </button>
              
              {videoUrl && (
                <div className={`flex items-center gap-3 px-3 py-1.5 rounded-xl border ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-gray-50 border-gray-200'}`}>
                  <div className="flex items-center gap-2">
                    <span className={`text-[9px] font-mono ${
                      !isRequirementMet('resolution')
                        ? 'text-red-500 font-bold' 
                        : (theme === 'dark' ? 'text-white/40' : 'text-gray-400')
                    }`}>
                      {videoMetadata?.resolution || '--'}
                    </span>
                    <div className={`w-px h-2 ${theme === 'dark' ? 'bg-white/10' : 'bg-gray-200'}`} />
                    <span className={`text-[9px] font-mono ${
                      !isRequirementMet('size')
                        ? 'text-red-500 font-bold' 
                        : (theme === 'dark' ? 'text-white/40' : 'text-gray-400')
                    }`}>
                      {videoMetadata?.size || '--'}
                    </span>
                    <div className={`w-px h-2 ${theme === 'dark' ? 'bg-white/10' : 'bg-gray-200'}`} />
                    <span className={`text-[9px] font-mono ${
                      !isRequirementMet('duration')
                        ? 'text-red-500 font-bold'
                        : 'text-[#F27D26]'
                    }`}>
                            <HeaderTime 
                              videoRef={videoRef}
                              duration={duration}
                              isRequirementMet={isRequirementMet}
                              theme={theme}
                            />
                    </span>
                  </div>
                </div>
              )}

              {videoUrl && (
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => setShowRequirementsModal(true)}
                    className={`flex items-center gap-2 px-3 py-2 rounded-lg text-[10px] font-bold uppercase transition-all border ${theme === 'dark' ? 'bg-blue-500/10 border-blue-500/30 text-blue-400' : 'bg-blue-50 border-blue-100 text-blue-600 hover:bg-blue-100'}`}
                  >
                    <ClipboardList size={14} />
                    审核要求
                  </button>
                  <button 
                    onClick={detectBlackFrames}
                    disabled={isDetecting}
                    className={`relative overflow-hidden flex items-center gap-2 px-4 py-2.5 rounded-xl text-[10px] font-bold uppercase transition-all border shadow-lg ${
                      isDetecting 
                        ? 'cursor-not-allowed border-purple-500/50' 
                        : 'hover:scale-105 active:scale-95 border-purple-500/30 hover:border-purple-500/60'
                    } ${
                      theme === 'dark' 
                        ? 'bg-purple-500/10 text-purple-400' 
                        : 'bg-purple-50 text-purple-600 hover:bg-purple-100'
                    }`}
                  >
                    {/* Progress Background Fill */}
                    {isDetecting && (
                      <motion.div 
                        className="absolute inset-0 bg-purple-500/20 origin-left"
                        initial={{ scaleX: 0 }}
                        animate={{ scaleX: detectionProgress / 100 }}
                        transition={{ type: "spring", bounce: 0, duration: 0.5 }}
                      />
                    )}
                    
                    <span className="relative z-10 flex items-center gap-2">
                      <Sparkles 
                        size={14} 
                        className={`${isDetecting ? 'animate-spin text-purple-400' : 'text-purple-500'}`} 
                      />
                      <span className="tracking-widest">
                        {isDetecting ? `检测中 ${detectionProgress}%` : '黑帧检测'}
                      </span>
                    </span>
                  </button>
                </div>
              )}
              {videoList.length > 1 && (
                <div className="flex items-center gap-2">
                  <button 
                    onClick={exportBatchReport}
                    className={`flex items-center gap-2 px-3 py-2 rounded-lg text-[10px] font-bold uppercase transition-all border ${theme === 'dark' ? 'bg-green-500/10 border-green-500/30 text-green-400' : 'bg-green-50 border-green-100 text-green-600 hover:bg-green-100'}`}
                    title="导出视频列表汇总"
                  >
                    <FileSpreadsheet size={14} />
                    汇总表
                  </button>
                  <button 
                    onClick={exportAllMarkersToExcel}
                    className={`flex items-center gap-2 px-3 py-2 rounded-lg text-[10px] font-bold uppercase transition-all border ${theme === 'dark' ? 'bg-blue-500/10 border-blue-500/30 text-blue-400' : 'bg-blue-50 border-blue-100 text-blue-600 hover:bg-blue-100'}`}
                    title="导出所有视频的详细批注意见"
                  >
                    <FileStack size={14} />
                    全量Excel
                  </button>
                  <button 
                    onClick={exportAllMarkersToWord}
                    className={`flex items-center gap-2 px-3 py-2 rounded-lg text-[10px] font-bold uppercase transition-all border ${theme === 'dark' ? 'bg-amber-500/10 border-amber-500/30 text-amber-400' : 'bg-amber-50 border-amber-100 text-amber-600 hover:bg-amber-100'}`}
                    title="导出所有视频的详细批注意见 (Word)"
                  >
                    <FileText size={14} />
                    全量Word
                  </button>
                </div>
              )}
              <button 
                onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
                className={`p-2 rounded-lg transition-all border ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 border-white/10 text-white' : 'bg-gray-100 hover:bg-gray-200 border-gray-200 text-gray-700'}`}
              >
                {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
              </button>
              <button 
                onClick={() => triggerVideoUpload(true)}
                className={`px-4 py-2 text-[10px] font-bold uppercase transition-all rounded-lg flex items-center gap-2 border ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 border-white/10 text-white' : 'bg-gray-100 hover:bg-gray-200 border-gray-200 text-gray-700'}`}
              >
                <Download size={14} />
                导入视频
              </button>
              <button 
                onClick={triggerFolderUpload}
                className={`px-4 py-2 text-[10px] font-bold uppercase transition-all rounded-lg flex items-center gap-2 border ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 border-white/10 text-white' : 'bg-gray-100 hover:bg-gray-200 border-gray-200 text-gray-700'}`}
              >
                <Folder size={14} />
                导入文件夹
              </button>
            </div>
          </motion.header>
        </div>

        <main 
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          className="flex-1 grid grid-cols-1 lg:grid-cols-12 gap-4 min-h-0 pt-20 relative"
        >
          {/* Hidden Inputs for Uploads */}
          <input 
            type="file" 
            ref={videoInputRef} 
            className="hidden" 
            accept="video/*,.mp4,.mov,.avi,.mkv,.webm" 
            multiple 
            onChange={(e) => handleVideoUploadWithWarning(e, false)} 
          />
          <input 
            type="file" 
            ref={replaceVideoInputRef} 
            className="hidden" 
            accept="video/*,.mp4,.mov,.avi,.mkv,.webm" 
            onChange={(e) => handleVideoUploadWithWarning(e, true)} 
          />
          <input 
            type="file" 
            ref={folderInputRef} 
            className="hidden" 
            // @ts-ignore
            webkitdirectory="" 
            // @ts-ignore
            directory="" 
            onChange={(e) => handleVideoUploadWithWarning(e, false)} 
          />
          {/* Drag and Drop Overlay */}
          <AnimatePresence>
            {isDragging && (
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="absolute inset-0 z-[100] bg-[#F27D26]/20 backdrop-blur-sm border-4 border-dashed border-[#F27D26] rounded-3xl m-2 flex flex-col items-center justify-center pointer-events-none"
              >
                <div className="bg-white dark:bg-[#111] p-8 rounded-full shadow-2xl mb-4">
                  <Download size={48} className="text-[#F27D26] animate-bounce" />
                </div>
                <h2 className="text-2xl font-bold dark:text-white mb-2">松开鼠标导入视频</h2>
                <p className="text-sm dark:text-white/60">支持单个视频或整个文件夹</p>
              </motion.div>
            )}
          </AnimatePresence>
          
          {/* Left: Player & Input Section */}
          <div className="lg:col-span-8 flex flex-col gap-3 min-h-0 overflow-hidden">
            {videoUrl && (
              <div className="flex items-center gap-2 px-1">
                <div className={`${theme === 'dark' ? 'bg-[#F27D26]/10' : 'bg-[#F27D26]/5'} p-1 rounded-lg`}>
                  <Film size={12} className="text-[#F27D26]" />
                </div>
                <h2 className={`text-xs font-bold tracking-tight truncate ${theme === 'dark' ? 'text-white/90' : 'text-[#1A1A1A]'}`}>
                  {videoMetadata?.name || '未命名视频'}
                </h2>
                <div className={`ml-2 px-1.5 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider ${theme === 'dark' ? 'bg-white/5 text-white/30' : 'bg-black/5 text-black/30'}`}>
                  正在审阅
                </div>
              </div>
            )}
            {/* Video Container */}
            <div className={`bg-black w-full aspect-video flex-shrink-0 overflow-hidden shadow-2xl border relative flex items-center justify-center group ${theme === 'dark' ? 'border-white/5' : 'border-black/5'}`}>
              {videoUrl ? (
                <>
                  <video 
                    ref={videoRef}
                    src={videoUrl}
                    onTimeUpdate={handleTimeUpdate}
                    onLoadedMetadata={handleLoadedMetadata}
                    onPlay={handlePlay}
                    className="w-full h-full object-cover rounded-none"
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
                    className={`absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 rounded-none ${drawingMode === 'none' ? 'pointer-events-none' : 'pointer-events-auto cursor-crosshair'}`}
                  />
                  <canvas
                    ref={tempCanvasRef}
                    className={`absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 pointer-events-none rounded-none ${drawingMode === 'none' ? 'hidden' : ''}`}
                  />

                  {/* Weak Reminder (Success Message) - Bottom Right of Video */}
                  <AnimatePresence>
                    {successMessage && (
                      <motion.div
                        initial={{ opacity: 0, scale: 0.8, x: 20 }}
                        animate={{ opacity: 1, scale: 1, x: 0 }}
                        exit={{ opacity: 0, scale: 0.8, x: 20 }}
                        className="absolute bottom-20 right-4 bg-[#F27D26] text-white px-4 py-2 rounded-xl text-[10px] font-bold shadow-2xl z-50 border border-white/20 backdrop-blur-md"
                      >
                        {successMessage}
                      </motion.div>
                    )}
                  </AnimatePresence>

                  {/* Floating Drawing Toolbar */}
                  <div className="absolute top-4 left-1/2 -translate-x-1/2 flex items-center gap-3 bg-black/80 backdrop-blur-xl border border-white/10 p-2 rounded-none opacity-0 group-hover:opacity-100 transition-all duration-300 z-30 shadow-2xl scale-90 group-hover:scale-100">
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

                  {/* Full-width Custom Control Bar */}
                  <VideoControls 
                    videoRef={videoRef}
                    duration={duration}
                    markers={markers}
                    onSeek={seekTo}
                    onSetHoveredMarker={setHoveredMarker}
                    isDetecting={isDetecting}
                    detectionProgress={detectionProgress}
                    theme={theme}
                  />
                </>
              ) : (
                <div className="flex flex-col items-center opacity-20">
                  <Play size={80} strokeWidth={1} />
                  <p className="mt-6 text-xs uppercase tracking-[0.3em]">等待视频载入...</p>
                </div>
              )}
            </div>

            {/* Video Info Bar removed as per user request */}

            {/* Input & Audio Section (Side by Side) */}
            <div className="flex gap-3 shrink-0">
              {/* Review Input Section */}
              <div className={`flex-1 flex flex-col border rounded-2xl overflow-hidden shadow-xl transition-colors duration-300 ${theme === 'dark' ? 'bg-[#111] border-white/5' : 'bg-white border-gray-200 shadow-sm'}`}>
                <div className={`p-2 px-4 border-b flex items-center justify-between ${theme === 'dark' ? 'bg-white/[0.02] border-white/5' : 'bg-gray-50 border-gray-200'}`}>
                  <h3 className="text-[9px] font-bold uppercase tracking-widest flex items-center gap-2 shrink-0">
                    <Sparkles size={12} className="text-[#F27D26]" />
                    添加审核意见
                  </h3>
                  <div className="flex flex-wrap gap-1 ml-4 overflow-x-auto no-scrollbar">
                    {(Object.keys(CATEGORY_COLORS) as Array<keyof typeof CATEGORY_COLORS>).map(cat => (
                      <button
                        key={cat}
                        onClick={() => setSelectedCategory(cat)}
                        className={`px-2 py-0.5 rounded-lg text-[7px] font-bold uppercase transition-all whitespace-nowrap ${
                          selectedCategory === cat 
                          ? (theme === 'dark' ? 'bg-white text-black' : 'bg-black text-white')
                          : (theme === 'dark' ? 'bg-white/5 text-white/40 hover:bg-white/10' : 'bg-gray-100 text-gray-500 hover:bg-gray-200')
                        }`}
                      >
                        {CATEGORY_COLORS[cat].label}
                      </button>
                    ))}
                  </div>
                </div>
                <div className="p-3">
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
                      className={`w-full border rounded-xl p-3 text-xs focus:outline-none focus:border-[#F27D26]/40 transition-all resize-none h-16 ${theme === 'dark' ? 'bg-black/40 border-white/5 text-white' : 'bg-white border-gray-200 text-gray-800'}`}
                    />
                    <button
                      onClick={() => handleAddReview()}
                      disabled={!videoUrl}
                      className="absolute bottom-3 right-3 bg-[#F27D26] text-white p-2 rounded-lg hover:scale-105 active:scale-95 transition-all disabled:opacity-20 shadow-lg shadow-[#F27D26]/20"
                    >
                      <Sparkles size={14} />
                    </button>
                  </div>
                </div>
              </div>

              {/* Audio Visualizer (Compact) */}
              {videoUrl && (
                <div className="w-20 flex flex-col relative">
                  <AudioVisualizer 
                    theme={theme}
                    analyserL={analyserLRef.current}
                    analyserR={analyserRRef.current}
                    onImbalance={handleImbalanceChange}
                  />
                  
                  <AnimatePresence>
                    {showImbalanceAlert && isAudioImbalanced && !isAlertDismissed && (
                      <motion.div
                        initial={{ opacity: 0, scale: 0.8, x: "-50%", y: 20 }}
                        animate={{ opacity: 1, scale: 1, x: "-50%", y: 0 }}
                        exit={{ opacity: 0, scale: 0.8, x: "-50%", y: 20 }}
                        className="fixed bottom-24 left-1/2 w-56 bg-red-600 text-white p-3 rounded-2xl shadow-2xl flex flex-col items-center gap-2 z-[999] border border-white/20"
                      >
                        <button 
                          onClick={dismissAlert}
                          className="absolute top-2 right-2 p-1 hover:bg-white/10 rounded-full transition-colors"
                        >
                          <Trash2 size={12} className="text-white/60 hover:text-white" />
                        </button>
                        <div className="flex items-center gap-2 pr-4">
                          <AlertCircle size={16} className="animate-pulse shrink-0" />
                          <span className="text-xs font-bold whitespace-nowrap tracking-wider">声道严重不平衡</span>
                        </div>
                        <p className="text-[9px] text-white/80 text-center leading-tight px-1">检测到左右声道音量差异过大，已自动为您添加批注。</p>
                        <div className="w-full h-1 bg-white/20 rounded-full overflow-hidden mt-1">
                          <motion.div 
                            className="h-full bg-white"
                            animate={{ width: ["0%", "100%"] }}
                            transition={{ duration: 1.5, repeat: Infinity, ease: "linear" }}
                          />
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              )}
            </div>
          </div>

          {/* Right Part of Layout (Columns 9-12) */}
          <div className="lg:col-span-4 grid grid-cols-2 gap-4 h-full min-h-0">
            {/* Middle: Sidebar Section (Review List & Video List) - Column 9-10 */}
            <div className="flex flex-col gap-4 min-h-0">
              {/* Review List Section */}
              <div className={`flex-1 flex flex-col border rounded-2xl overflow-hidden shadow-xl min-h-0 transition-colors duration-300 ${theme === 'dark' ? 'bg-[#111] border-white/5' : 'bg-white border-gray-200 shadow-sm'}`}>
                <div className={`p-3 px-4 border-b flex items-center justify-between ${theme === 'dark' ? 'bg-white/[0.02] border-white/5' : 'bg-gray-50 border-gray-200'}`}>
                  <h3 className="text-[9px] font-bold uppercase tracking-widest flex items-center gap-2">
                    <Clock size={12} className="text-[#F27D26]" />
                    批注列表 ({markers.length})
                  </h3>
                  <button onClick={() => setMarkers([])} className={`p-1 rounded-lg transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/20' : 'hover:bg-black/5 text-black/20'} hover:text-red-500`}>
                    <Trash2 size={10} />
                  </button>
                </div>

                <div className="flex-1 overflow-y-auto custom-scrollbar p-2 space-y-2">
                  {markers.length === 0 ? (
                    <div className="h-full flex flex-col items-center justify-center opacity-10 text-center px-4">
                      <FileCode size={24} strokeWidth={1} />
                      <p className="mt-2 text-[7px] uppercase tracking-widest leading-relaxed">暂无批注</p>
                    </div>
                  ) : (
                    markers.map((m, i) => (
                      <MarkerItem 
                        key={`${m.seconds}-${i}`}
                        marker={m}
                        index={i}
                        theme={theme}
                        onSeek={seekTo}
                        onDelete={deleteMarker}
                      />
                    ))
                  )}
                </div>

                <div className="p-2 border-t flex gap-2">
                  <button 
                    onClick={exportReviewCommentsToExcel}
                    disabled={markers.length === 0}
                    className="flex-1 py-2 rounded-xl bg-green-600 text-white text-[8px] font-bold uppercase hover:scale-[1.02] active:scale-[0.98] transition-all disabled:opacity-20 flex items-center justify-center gap-1 shadow-lg shadow-green-600/20"
                    title="导出 Excel"
                  >
                    <FileSpreadsheet size={12} />
                    Excel
                  </button>
                  <button 
                    onClick={exportReviewCommentsToWord}
                    disabled={markers.length === 0}
                    className="flex-1 py-2 rounded-xl bg-blue-600 text-white text-[8px] font-bold uppercase hover:scale-[1.02] active:scale-[0.98] transition-all disabled:opacity-20 flex items-center justify-center gap-1 shadow-lg shadow-blue-600/20"
                    title="导出 Word"
                  >
                    <FileText size={12} />
                    Word
                  </button>
                </div>
              </div>

              {/* Video List Section (Bottom of Sidebar) */}
              {videoList.length > 0 && (
                <div className={`h-[180px] flex flex-col border rounded-2xl overflow-hidden shadow-xl transition-colors duration-300 ${theme === 'dark' ? 'bg-[#111] border-white/5' : 'bg-white border-gray-200 shadow-sm'}`}>
                  <div className={`p-3 px-4 border-b flex items-center justify-between ${theme === 'dark' ? 'bg-white/[0.02] border-white/5' : 'bg-gray-50 border-gray-200'}`}>
                    <h3 className="text-[9px] font-bold uppercase tracking-widest flex items-center gap-2">
                      <Film size={12} className="text-[#F27D26]" />
                      视频列表 ({videoList.length})
                    </h3>
                    <button 
                      onClick={() => { setVideoList([]); setVideoUrl(null); setSelectedVideoId(null); setMarkers([]); setVideoMetadata(null); }}
                      className={`p-1 rounded-lg transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/20' : 'hover:bg-black/5 text-black/20'} hover:text-red-500`}
                    >
                      <Trash2 size={10} />
                    </button>
                  </div>
                  <div className="flex-1 overflow-y-auto custom-scrollbar p-1.5 space-y-0.5">
                    {videoList.map((v) => (
                      <VideoListItem 
                        key={v.id}
                        video={v}
                        isSelected={selectedVideoId === v.id}
                        theme={theme}
                        onSelect={selectVideo}
                        onDelete={handleDeleteVideo}
                      />
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* Right: Quick Replies Sidebar - Column 11-12 */}
            <div className="flex flex-col gap-4 min-h-0">
              {/* Quick Replies Section */}
              <div className={`flex-1 flex flex-col border rounded-2xl overflow-hidden shadow-xl min-h-0 transition-colors duration-300 ${theme === 'dark' ? 'bg-[#111] border-white/5' : 'bg-white border-gray-200 shadow-sm'}`}>
                <div className={`p-3 px-4 border-b flex flex-col gap-2 ${theme === 'dark' ? 'bg-white/[0.02] border-white/5' : 'bg-gray-50 border-gray-200'}`}>
                  <div className="flex items-center justify-between">
                    <h3 className="text-[9px] font-bold uppercase tracking-widest flex items-center gap-2">
                      <Quote size={12} className="text-[#F27D26]" />
                      快捷话术
                    </h3>
                    <div className="flex items-center gap-1">
                      <button 
                        onClick={addCategory}
                        className={`p-1 rounded-lg transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/20' : 'hover:bg-black/5 text-black/20'} hover:text-[#F27D26]`}
                        title="新建分类"
                      >
                        <Plus size={10} />
                      </button>
                      <button 
                        onClick={handleImportQuickReplies} 
                        className={`p-1 rounded-lg transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/20' : 'hover:bg-black/5 text-black/20'} hover:text-[#F27D26]`}
                        title="导入 Excel/JSON"
                      >
                        <FileUp size={10} />
                      </button>
                      <button 
                        onClick={handleExportQuickReplies} 
                        className={`p-1 rounded-lg transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/20' : 'hover:bg-black/5 text-black/20'} hover:text-[#F27D26]`}
                        title="导出 Excel"
                      >
                        <Download size={10} />
                      </button>
                    </div>
                  </div>
                  
                  {/* Category Tabs */}
                  <div className="flex overflow-x-auto no-scrollbar gap-1 py-1 group/tabs">
                    {Object.keys(quickReplies).map(cat => (
                      <div key={cat} className="relative shrink-0 flex items-center">
                        <button
                          onClick={() => setActiveQuickReplyTab(cat)}
                          onContextMenu={(e) => { e.preventDefault(); renameCategory(cat); }}
                          className={`px-2 py-1 rounded-md text-[7px] font-bold uppercase transition-all whitespace-nowrap border ${
                            activeQuickReplyTab === cat 
                            ? (theme === 'dark' ? 'bg-[#F27D26] border-[#F27D26] text-white' : 'bg-[#F27D26] border-[#F27D26] text-white shadow-sm')
                            : (theme === 'dark' ? 'bg-white/5 border-white/10 text-white/40 hover:bg-white/10' : 'bg-white border-gray-200 text-gray-500 hover:bg-gray-50')
                          }`}
                          title={`右键重命名 ${cat}`}
                        >
                          {cat}
                        </button>
                        <button 
                          onClick={(e) => { e.stopPropagation(); deleteCategory(cat); }}
                          className="absolute -top-1 -right-1 bg-red-500 text-white rounded-full p-0.5 opacity-0 group-hover/tabs:opacity-100 transition-opacity hover:scale-110 z-10"
                        >
                          <X size={6} />
                        </button>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="flex-1 overflow-y-auto custom-scrollbar p-2 space-y-1.5 focus-within:ring-2 focus-within:ring-[#F27D26]/20 transition-all">
                  {(!quickReplies[activeQuickReplyTab] || quickReplies[activeQuickReplyTab].length === 0) ? (
                    <div className="h-full flex flex-col items-center justify-center opacity-10 text-center px-4">
                      <Quote size={20} strokeWidth={1} />
                      <p className="mt-2 text-[7px] uppercase tracking-widest leading-relaxed">
                        {activeQuickReplyTab ? '该分类下无话术' : '请先选择分类或导入'}
                      </p>
                    </div>
                  ) : (
                    quickReplies[activeQuickReplyTab].map((reply, i) => (
                      <QuickReplyItem 
                        key={`${activeQuickReplyTab}-${i}`}
                        reply={reply}
                        index={i}
                        tab={activeQuickReplyTab}
                        theme={theme}
                        onRemove={removeQuickReply}
                        onDoubleClick={handleAddQuickReplyToComment}
                      />
                    ))
                  )}
                </div>
                
                <div className={`p-2 border-t text-[7px] opacity-40 text-center flex items-center justify-center gap-1 ${theme === 'dark' ? 'bg-white/[0.01]' : 'bg-gray-50'}`}>
                  <Info size={8} /> 分层自定义，双击触发批注
                </div>
              </div>
            </div>
          </div>
        </main>

        {/* Category Modal Overlay */}
        <AnimatePresence>
          {categoryModal && (
            <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setCategoryModal(null)}
                className="absolute inset-0 bg-black/60 backdrop-blur-sm"
              />
              <motion.div 
                initial={{ opacity: 0, scale: 0.9, y: 20 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.9, y: 20 }}
                className={`relative w-full max-w-sm rounded-3xl shadow-2xl overflow-hidden border ${theme === 'dark' ? 'bg-[#111] border-white/10' : 'bg-white border-gray-200'}`}
              >
                <div className="p-6">
                  <h3 className="text-sm font-bold uppercase tracking-widest mb-4">
                    {categoryModal.type === 'add' ? '新增分类' : categoryModal.type === 'rename' ? '重命名分类' : '删除确认'}
                  </h3>
                  
                  {categoryModal.type !== 'delete' ? (
                    <input
                      autoFocus
                      type="text"
                      value={categoryInput}
                      onChange={(e) => setCategoryInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') categoryModal.type === 'add' ? confirmAddCategory() : confirmRenameCategory();
                        if (e.key === 'Escape') setCategoryModal(null);
                      }}
                      placeholder="输入分类名称..."
                      className={`w-full px-4 py-3 rounded-xl text-sm border focus:ring-2 focus:ring-[#F27D26] outline-none transition-all ${theme === 'dark' ? 'bg-white/5 border-white/10 text-white' : 'bg-gray-50 border-gray-200'}`}
                    />
                  ) : (
                    <p className={`text-sm ${theme === 'dark' ? 'text-white/60' : 'text-gray-500'}`}>
                      确定要删除分类 <span className="text-[#F27D26] font-bold">"{categoryModal.target}"</span> 及其所有话术吗？此操作不可撤销。
                    </p>
                  )}
                  
                  <div className="flex gap-2 mt-6">
                    <button 
                      onClick={() => setCategoryModal(null)}
                      className={`flex-1 py-3 rounded-xl text-xs font-bold uppercase tracking-widest transition-all ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 text-white/40' : 'bg-gray-100 hover:bg-gray-200 text-gray-500'}`}
                    >
                      取消
                    </button>
                    <button 
                      onClick={() => {
                        if (categoryModal.type === 'add') confirmAddCategory();
                        else if (categoryModal.type === 'rename') confirmRenameCategory();
                        else confirmDeleteCategory();
                      }}
                      className={`flex-1 py-3 rounded-xl text-xs font-bold uppercase tracking-widest text-white transition-all shadow-lg ${categoryModal.type === 'delete' ? 'bg-red-500 shadow-red-500/20' : 'bg-[#F27D26] shadow-[#F27D26]/20'}`}
                    >
                      {categoryModal.type === 'delete' ? '确定删除' : '保存'}
                    </button>
                  </div>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Hover Tooltip for Timeline */}
        <AnimatePresence>
          {hoveredMarker && (
            <motion.div 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 10 }}
              className={`fixed bottom-32 left-1/2 -translate-x-1/2 border px-4 py-2 rounded-lg shadow-2xl z-50 pointer-events-none ${theme === 'dark' ? 'bg-[#1A1A1A] border-white/10 text-white' : 'bg-white border-gray-200 text-gray-900'}`}
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
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>Space / K</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">后退 5s</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>J / ←</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">前进 5s</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>L / →</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">上一帧</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>D / ,</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">下一帧</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>F / .</kbd>
                    </div>
                  </div>
                  <div className="space-y-2">
                    <p className="text-[10px] font-bold opacity-30 uppercase tracking-wider mb-2">审核功能</p>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">聚焦输入框</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>M</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">发送批注</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>Ctrl+Enter</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">选择分类</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>1 - 5</kbd>
                    </div>
                    <div className="flex justify-between items-center text-xs">
                      <span className="opacity-50">全屏播放</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>Enter</kbd>
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
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>C / X</kbd>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="opacity-50">切换倍速</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>Z</kbd>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="opacity-50">音量调节</span>
                      <kbd className={`px-1.5 py-0.5 rounded font-mono text-[10px] ${theme === 'dark' ? 'bg-white/5 border-white/10' : 'bg-black/5 border-black/10'}`}>↑ / ↓</kbd>
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
      {/* Requirements Modal */}
      <AnimatePresence>
        {showRequirementsModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className={`w-full max-w-md rounded-3xl shadow-2xl overflow-hidden border ${theme === 'dark' ? 'bg-[#111] border-white/10' : 'bg-white border-gray-200'}`}
            >
              <div className={`p-6 border-b ${theme === 'dark' ? 'border-white/5' : 'border-gray-100'}`}>
                <div className="flex items-center justify-between">
                  <h3 className={`text-lg font-bold flex items-center gap-2 ${theme === 'dark' ? 'text-white' : 'text-gray-900'}`}>
                    <ClipboardList className="text-[#F27D26]" size={20} />
                    审核要求设置
                  </h3>
                  <button onClick={() => setShowRequirementsModal(false)} className={`p-2 rounded-xl transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/40' : 'hover:bg-gray-100 text-gray-400'}`}>
                    <Trash2 size={18} />
                  </button>
                </div>
              </div>
              
              <div className="p-6 space-y-6">
                <div className="space-y-2 relative">
                  <label className={`text-[10px] font-bold uppercase tracking-widest ${theme === 'dark' ? 'text-white/40' : 'text-gray-400'}`}>分辨率要求</label>
                  <div className="relative">
                    <button 
                      onClick={() => setIsResolutionDropdownOpen(!isResolutionDropdownOpen)}
                      className={`w-full px-4 py-3 rounded-2xl border text-sm flex items-center justify-between transition-all ${theme === 'dark' ? 'bg-white/5 border-white/10 text-white' : 'bg-gray-50 border-gray-200 text-gray-900'}`}
                    >
                      <span>{requirements.resolution || '不限制'}</span>
                      <ChevronDown size={16} className={`transition-transform duration-300 ${isResolutionDropdownOpen ? 'rotate-180' : ''}`} />
                    </button>
                    
                    <AnimatePresence>
                      {isResolutionDropdownOpen && (
                        <>
                          <div className="fixed inset-0 z-[110]" onClick={() => setIsResolutionDropdownOpen(false)} />
                          <motion.div 
                            initial={{ opacity: 0, y: -10, scale: 0.95 }}
                            animate={{ opacity: 1, y: 4, scale: 1 }}
                            exit={{ opacity: 0, y: -10, scale: 0.95 }}
                            className={`absolute left-0 right-0 z-[120] rounded-2xl border shadow-2xl overflow-hidden ${theme === 'dark' ? 'bg-[#1a1a1a] border-white/10' : 'bg-white border-gray-200'}`}
                          >
                            {[
                              { label: '不限制', value: '' },
                              { label: '1920 x 1080 (1080p)', value: '1920 x 1080' },
                              { label: '1280 x 720 (720p)', value: '1280 x 720' },
                              { label: '3840 x 2160 (4K)', value: '3840 x 2160' },
                              { label: '1080 x 1920 (竖屏 1080p)', value: '1080 x 1920' },
                              { label: '720 x 1280 (竖屏 720p)', value: '720 x 1280' }
                            ].map((opt) => (
                              <button
                                key={opt.value}
                                onClick={() => {
                                  setRequirements(prev => ({ ...prev, resolution: opt.value }));
                                  setIsResolutionDropdownOpen(false);
                                }}
                                className={`w-full px-4 py-3 text-left text-sm transition-colors ${
                                  requirements.resolution === opt.value 
                                    ? 'bg-[#F27D26] text-white' 
                                    : (theme === 'dark' ? 'hover:bg-white/5 text-white' : 'hover:bg-gray-50 text-gray-700')
                                }`}
                              >
                                {opt.label}
                              </button>
                            ))}
                          </motion.div>
                        </>
                      )}
                    </AnimatePresence>
                  </div>
                </div>
                
                <div className="space-y-2">
                  <label className={`text-[10px] font-bold uppercase tracking-widest ${theme === 'dark' ? 'text-white/40' : 'text-gray-400'}`}>最大时长要求 (分钟)</label>
                  <input 
                    type="number" 
                    step="0.1"
                    value={requirements.duration || ''}
                    onChange={(e) => setRequirements(prev => ({ ...prev, duration: e.target.value ? parseFloat(e.target.value) : null }))}
                    placeholder="不限制"
                    className={`w-full px-4 py-3 rounded-2xl border text-sm focus:outline-none focus:border-[#F27D26] transition-all ${theme === 'dark' ? 'bg-white/5 border-white/10 text-white' : 'bg-gray-50 border-gray-200 text-gray-900'}`}
                  />
                </div>
                
                <div className="space-y-2">
                  <label className={`text-[10px] font-bold uppercase tracking-widest ${theme === 'dark' ? 'text-white/40' : 'text-gray-400'}`}>最大大小要求 (MB)</label>
                  <input 
                    type="number" 
                    value={requirements.size || ''}
                    onChange={(e) => setRequirements(prev => ({ ...prev, size: e.target.value ? parseFloat(e.target.value) : null }))}
                    placeholder="不限制"
                    className={`w-full px-4 py-3 rounded-2xl border text-sm focus:outline-none focus:border-[#F27D26] transition-all ${theme === 'dark' ? 'bg-white/5 border-white/10 text-white' : 'bg-gray-50 border-gray-200 text-gray-900'}`}
                  />
                </div>
              </div>
              
              <div className={`p-6 border-t flex gap-3 ${theme === 'dark' ? 'border-white/5 bg-white/[0.02]' : 'border-black/5 bg-black/[0.02]'}`}>
                <button 
                  onClick={() => setShowRequirementsModal(false)}
                  className={`flex-1 py-3 rounded-2xl font-bold text-sm transition-all ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 text-white' : 'bg-black/5 hover:bg-black/10 text-black'}`}
                >
                  取消
                </button>
                <button 
                  onClick={() => {
                    setShowRequirementsModal(false);
                    setSuccessMessage('审核要求已保存');
                    setTimeout(() => setSuccessMessage(null), 2000);
                  }}
                  className="flex-1 py-3 rounded-2xl bg-[#F27D26] hover:bg-[#D96A1D] text-white font-bold text-sm transition-all shadow-lg shadow-[#F27D26]/20"
                >
                  保存设置
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
        {/* Unsaved Changes Warning Modal */}
        <AnimatePresence>
          {showUnsavedWarning && (
            <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4">
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setShowUnsavedWarning(false)}
                className="absolute inset-0 bg-black/60 backdrop-blur-sm"
              />
              <motion.div 
                initial={{ opacity: 0, scale: 0.9, y: 20 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.9, y: 20 }}
                className={`relative w-full max-w-md rounded-3xl shadow-2xl overflow-hidden border ${theme === 'dark' ? 'bg-[#111] border-white/10' : 'bg-white border-gray-200'}`}
              >
                <div className="p-8 text-center space-y-6">
                  <div className="w-16 h-16 bg-amber-500/10 rounded-full flex items-center justify-center mx-auto">
                    <AlertCircle size={32} className="text-amber-500" />
                  </div>
                  <div className="space-y-2">
                    <h3 className={`text-xl font-bold ${theme === 'dark' ? 'text-white' : 'text-gray-900'}`}>未保存的审核意见</h3>
                    <p className={`text-sm ${theme === 'dark' ? 'text-white/60' : 'text-gray-500'}`}>
                      检测到当前视频已有批注。如果重新导入视频，当前的审核意见将会丢失。建议先导出汇总表。
                    </p>
                  </div>
                  <div className="grid grid-cols-1 gap-3 pt-4">
                    <button 
                      onClick={() => {
                        exportReviewCommentsToWord();
                        setShowUnsavedWarning(false);
                      }}
                      className="w-full py-4 rounded-2xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm transition-all shadow-lg shadow-blue-600/20 flex items-center justify-center gap-2"
                    >
                      <FileText size={18} />
                      导出汇总表 (Word)
                    </button>
                    <button 
                      onClick={confirmDiscardAndUpload}
                      className={`w-full py-4 rounded-2xl font-bold text-sm transition-all ${theme === 'dark' ? 'bg-red-500/10 hover:bg-red-500/20 text-red-400' : 'bg-red-50 hover:bg-red-100 text-red-600'}`}
                    >
                      确认不保留，直接导入
                    </button>
                    <button 
                      onClick={() => setShowUnsavedWarning(false)}
                      className={`w-full py-4 rounded-2xl font-bold text-sm transition-all ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 text-white' : 'bg-gray-100 hover:bg-gray-200 text-gray-700'}`}
                    >
                      取消
                    </button>
                  </div>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

      {/* Duration Verification Modal */}
      <AnimatePresence>
        {showDurationCheck && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className={`w-full max-w-3xl rounded-3xl shadow-2xl overflow-hidden border flex flex-col max-h-[90vh] ${theme === 'dark' ? 'bg-[#111] border-white/10' : 'bg-white border-gray-200'}`}
            >
              <div className={`p-6 border-b shrink-0 ${theme === 'dark' ? 'border-white/5' : 'border-gray-100'}`}>
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-4">
                    <h3 className={`text-lg font-bold flex items-center gap-2 ${theme === 'dark' ? 'text-white' : 'text-gray-900'}`}>
                      <FileStack className="text-[#F27D26]" size={20} />
                      表格时长核对
                    </h3>
                    <div className={`flex p-1 rounded-xl ${theme === 'dark' ? 'bg-white/5' : 'bg-gray-100'}`}>
                      <button 
                        onClick={() => setShowHistory(false)}
                        className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${!showHistory ? (theme === 'dark' ? 'bg-white/10 text-white' : 'bg-white text-gray-900 shadow-sm') : 'text-gray-500 hover:text-gray-700'}`}
                      >
                        当前核对
                      </button>
                      <button 
                        onClick={() => setShowHistory(true)}
                        className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${showHistory ? (theme === 'dark' ? 'bg-white/10 text-white' : 'bg-white text-gray-900 shadow-sm') : 'text-gray-500 hover:text-gray-700'}`}
                      >
                        历史记录
                      </button>
                    </div>
                  </div>
                  <div className="flex items-center gap-2">
                    {showHistory && verificationHistory.length > 0 && (
                      <button 
                        onClick={() => {
                          if (window.confirm('确定要清空所有历史记录吗？')) {
                            setVerificationHistory([]);
                            localStorage.removeItem('videoVerificationHistory');
                          }
                        }}
                        className={`p-2 rounded-xl transition-colors ${theme === 'dark' ? 'hover:bg-red-500/10 text-red-400' : 'hover:bg-red-50 text-red-500'}`}
                        title="清空历史"
                      >
                        <Trash2 size={18} />
                      </button>
                    )}
                    {!showHistory && (
                      <button 
                        onClick={() => { 
                          setVerificationRows([{ name: '', expected: '' }]); 
                          setVerificationResults([]); 
                        }} 
                        className={`p-2 rounded-xl transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/40' : 'hover:bg-gray-100 text-gray-400'}`}
                        title="清空当前"
                      >
                        <Trash2 size={18} />
                      </button>
                    )}
                    <button 
                      onClick={() => { setShowDurationCheck(false); }} 
                      className={`p-2 rounded-xl transition-colors ${theme === 'dark' ? 'hover:bg-white/5 text-white/40' : 'hover:bg-gray-100 text-gray-400'}`}
                    >
                      <X size={18} />
                    </button>
                  </div>
                </div>
              </div>
              
              <div className="p-6 space-y-6 overflow-y-auto flex-1">
                {!showHistory ? (
                  <div className="space-y-6">
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <label className={`text-[10px] font-bold uppercase tracking-widest ${theme === 'dark' ? 'text-white/40' : 'text-gray-400'}`}>核对单元格</label>
                        <span className="text-[10px] text-[#F27D26] font-medium">提示：支持直接粘贴整列数据</span>
                      </div>
                      <div className={`rounded-2xl border overflow-hidden ${theme === 'dark' ? 'border-white/10' : 'border-gray-200'}`}>
                        <table className="w-full border-collapse">
                          <thead>
                            <tr className={`${theme === 'dark' ? 'bg-white/5' : 'bg-gray-50'}`}>
                              <th className={`w-12 py-3 text-[10px] font-bold uppercase ${theme === 'dark' ? 'text-white/40 border-white/5' : 'text-gray-400 border-gray-100'} border-b border-r`}>#</th>
                              <th className={`py-3 px-4 text-left text-[10px] font-bold uppercase ${theme === 'dark' ? 'text-white/40 border-white/5' : 'text-gray-400 border-gray-100'} border-b border-r`}>视频名称 (粘贴列)</th>
                              <th className={`py-3 px-4 text-left text-[10px] font-bold uppercase ${theme === 'dark' ? 'text-white/40 border-white/5' : 'text-gray-400 border-gray-100'} border-b border-r`}>登记时长 (粘贴列)</th>
                              <th className={`w-12 py-3 text-[10px] font-bold uppercase ${theme === 'dark' ? 'text-white/40 border-white/5' : 'text-gray-400 border-gray-100'} border-b`}>操作</th>
                            </tr>
                          </thead>
                          <tbody>
                            {verificationRows.map((row, i) => (
                              <tr key={i} className={`${theme === 'dark' ? 'hover:bg-white/[0.02]' : 'hover:bg-gray-50/50'} transition-colors`}>
                                <td className={`py-2 text-center text-[10px] font-mono ${theme === 'dark' ? 'text-white/20 border-white/5' : 'text-gray-300 border-gray-100'} border-b border-r`}>{i + 1}</td>
                                <td className={`py-1 px-2 border-b border-r ${theme === 'dark' ? 'border-white/5' : 'border-gray-100'}`}>
                                  <input 
                                    type="text"
                                    value={row.name}
                                    onChange={(e) => updateRow(i, 'name', e.target.value)}
                                    onPaste={(e) => {
                                      const text = e.clipboardData.getData('text');
                                      if (text.includes('\t') || (text.includes('\n') && text.includes('  '))) {
                                        e.preventDefault();
                                        handlePasteTable(text);
                                      } else if (text.includes('\n')) {
                                        e.preventDefault();
                                        handlePasteColumn('name', text);
                                      }
                                    }}
                                    placeholder="粘贴或输入名称"
                                    className={`w-full px-3 py-2 rounded-xl text-sm focus:outline-none transition-all ${theme === 'dark' ? 'bg-transparent text-white placeholder:text-white/10' : 'bg-transparent text-gray-900 placeholder:text-gray-300'}`}
                                  />
                                </td>
                                <td className={`py-1 px-2 border-b border-r ${theme === 'dark' ? 'border-white/5' : 'border-gray-100'}`}>
                                  <input 
                                    type="text"
                                    value={row.expected}
                                    onChange={(e) => updateRow(i, 'expected', e.target.value)}
                                    onPaste={(e) => {
                                      const text = e.clipboardData.getData('text');
                                      if (text.includes('\n')) {
                                        e.preventDefault();
                                        handlePasteColumn('expected', text);
                                      }
                                    }}
                                    placeholder="粘贴或输入时长"
                                    className={`w-full px-3 py-2 rounded-xl text-sm font-mono focus:outline-none transition-all ${theme === 'dark' ? 'bg-transparent text-white placeholder:text-white/10' : 'bg-transparent text-gray-900 placeholder:text-gray-300'}`}
                                  />
                                </td>
                                <td className={`py-2 text-center border-b ${theme === 'dark' ? 'border-white/5' : 'border-gray-100'}`}>
                                  <button 
                                    onClick={() => removeRow(i)}
                                    className={`p-1.5 rounded-lg transition-colors ${theme === 'dark' ? 'hover:bg-red-500/10 text-white/20 hover:text-red-500' : 'hover:bg-red-50 text-gray-300 hover:text-red-500'}`}
                                  >
                                    <X size={14} />
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <button 
                        onClick={() => setVerificationRows(prev => [...prev, { name: '', expected: '' }])}
                        className={`w-full py-2 rounded-xl border-2 border-dashed text-[10px] font-bold uppercase tracking-widest transition-all ${theme === 'dark' ? 'border-white/5 text-white/20 hover:border-white/10 hover:text-white/40' : 'border-gray-100 text-gray-300 hover:border-gray-200 hover:text-gray-400'}`}
                      >
                        + 添加一行
                      </button>
                    </div>

                    <button 
                      onClick={handleVerifyDurations}
                      disabled={verificationRows.every(r => !r.name && !r.expected)}
                      className="w-full py-4 rounded-2xl bg-[#F27D26] hover:bg-[#D96A1D] text-white font-bold text-sm transition-all shadow-lg shadow-[#F27D26]/20 disabled:opacity-50 flex items-center justify-center gap-2"
                    >
                      <CheckCircle size={18} />
                      立即核对
                    </button>

                    {verificationResults.length > 0 && (
                      <div className="space-y-3">
                        <label className={`text-[10px] font-bold uppercase tracking-widest ${theme === 'dark' ? 'text-white/40' : 'text-gray-400'}`}>核对结果</label>
                        <div className="grid grid-cols-1 gap-2">
                          {verificationResults.map((res, i) => (
                            <div key={i} className={`p-4 rounded-2xl border flex items-center justify-between gap-4 ${theme === 'dark' ? 'bg-white/5 border-white/5' : 'bg-gray-50 border-gray-100'}`}>
                              <div className="min-w-0 flex-1">
                                <p className={`text-xs font-bold truncate ${theme === 'dark' ? 'text-white' : 'text-gray-900'}`}>{res.name}</p>
                                <p className="text-[10px] opacity-50 mt-1">表格登记: {res.expected}</p>
                              </div>
                              <div className="text-right shrink-0">
                                <div className="flex items-center gap-2 justify-end">
                                  <span className={`text-xs font-mono font-bold ${
                                    res.status === 'match' ? 'text-green-500' : 
                                    res.status === 'mismatch' ? 'text-red-500' : 'text-gray-400'
                                  }`}>
                                    {res.actual}
                                  </span>
                                  {res.status === 'match' && <CheckCircle size={14} className="text-green-500" />}
                                  {res.status === 'mismatch' && <AlertCircle size={14} className="text-red-500" />}
                                  {res.status === 'not_found' && <Info size={14} className="text-gray-400" />}
                                </div>
                                <p className={`text-[9px] mt-1 font-bold uppercase ${
                                  res.status === 'match' ? 'text-green-500/50' : 
                                  res.status === 'mismatch' ? 'text-red-500/50' : 'text-gray-400/50'
                                }`}>
                                  {res.status === 'match' ? '时长一致' : 
                                   res.status === 'mismatch' ? '时长不符' : '未找到视频'}
                                </p>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                ) : (
                  <div className="space-y-4">
                    {verificationHistory.length === 0 ? (
                      <div className="py-20 text-center space-y-4">
                        <div className={`w-16 h-16 rounded-full flex items-center justify-center mx-auto ${theme === 'dark' ? 'bg-white/5' : 'bg-gray-50'}`}>
                          <History className="text-gray-400" size={32} />
                        </div>
                        <p className="text-sm text-gray-500">暂无核对历史记录</p>
                      </div>
                    ) : (
                      verificationHistory.map((item) => (
                        <div key={item.id} className={`rounded-2xl border overflow-hidden ${theme === 'dark' ? 'border-white/10 bg-white/[0.02]' : 'border-gray-100 bg-gray-50/30'}`}>
                          <div className={`px-4 py-3 border-b flex items-center justify-between ${theme === 'dark' ? 'border-white/5 bg-white/5' : 'border-gray-100 bg-gray-50'}`}>
                            <div className="flex items-center gap-3">
                              <span className={`text-[10px] font-mono ${theme === 'dark' ? 'text-white/40' : 'text-gray-400'}`}>
                                {new Date(item.timestamp).toLocaleString()}
                              </span>
                              <span className={`px-2 py-0.5 rounded-full text-[9px] font-bold uppercase ${theme === 'dark' ? 'bg-white/10 text-white/60' : 'bg-white text-gray-400 shadow-sm'}`}>
                                {item.results.length} 个视频
                              </span>
                            </div>
                            <div className="flex items-center gap-3">
                              <button 
                                onClick={() => {
                                  setVerificationRows(item.results.map(r => ({ name: r.name, expected: r.expected })));
                                  setVerificationResults(item.results);
                                  setShowHistory(false);
                                }}
                                className="text-[10px] font-bold text-[#F27D26] hover:underline"
                              >
                                恢复此记录
                              </button>
                              <button 
                                onClick={() => removeHistoryItem(item.id)}
                                className={`p-1.5 rounded-lg transition-colors ${theme === 'dark' ? 'hover:bg-red-500/10 text-white/20 hover:text-red-500' : 'hover:bg-red-50 text-gray-300 hover:text-red-500'}`}
                                title="删除此记录"
                              >
                                <Trash2 size={12} />
                              </button>
                            </div>
                          </div>
                          <div className="p-3 grid grid-cols-1 gap-2">
                            {item.results.slice(0, 3).map((res, idx) => (
                              <div key={idx} className="flex items-center justify-between text-[11px]">
                                <span className={`truncate flex-1 ${theme === 'dark' ? 'text-white/60' : 'text-gray-600'}`}>{res.name}</span>
                                <div className="flex items-center gap-2 ml-4">
                                  <span className="text-gray-400 font-mono">{res.expected}</span>
                                  <span className={res.status === 'match' ? 'text-green-500' : res.status === 'mismatch' ? 'text-red-500' : 'text-gray-400'}>
                                    {res.status === 'match' ? '✓' : res.status === 'mismatch' ? '✗' : '?'}
                                  </span>
                                </div>
                              </div>
                            ))}
                            {item.results.length > 3 && (
                              <p className="text-[10px] text-gray-400 text-center mt-1">... 还有 {item.results.length - 3} 个视频</p>
                            )}
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                )}
              </div>
              
              <div className={`p-6 border-t shrink-0 ${theme === 'dark' ? 'border-white/5 bg-white/[0.02]' : 'border-black/5 bg-black/[0.02]'}`}>
                <button 
                  onClick={() => { setShowDurationCheck(false); }}
                  className={`w-full py-4 rounded-2xl font-bold text-sm transition-all ${theme === 'dark' ? 'bg-white/5 hover:bg-white/10 text-white' : 'bg-black/5 hover:bg-black/10 text-black'}`}
                >
                  关闭窗口
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

