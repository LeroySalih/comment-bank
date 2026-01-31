'use client';

import { useState, useRef, useEffect } from 'react';
import { createPortal } from 'react-dom';

interface TooltipProps {
  children: React.ReactNode;
  content: React.ReactNode;
  maxWidth?: string;
  className?: string;
}

export default function Tooltip({ children, content, maxWidth = "max-w-xs", className = "" }: TooltipProps) {
  const [visible, setVisible] = useState(false);
  const [pos, setPos] = useState({ top: 0, left: 0 });
  const triggerRef = useRef<HTMLDivElement>(null);

  const calculatePosition = () => {
    if (triggerRef.current) {
      const rect = triggerRef.current.getBoundingClientRect();
      // Basic positioning: Bottom Left aligned
      let top = rect.bottom + 8; // 8px gap
      let left = rect.left;

      // Check if it goes off bottom screen
      if (top + 100 > window.innerHeight) {
         // Flip to top
         top = rect.top - 8 - 50; // Approximation, better would be to measure tooltip height
         // Without ref to tooltip content, we guess or just flip. 
         // Let's stick to bottom unless very close to edge, then maybe adjust.
         // For now, simpler is reliable: just render.
      }
      
      // Check right edge
      if (left + 300 > window.innerWidth) {
          left = window.innerWidth - 320; // Shift left
      }

      setPos({ top, left });
    }
  }

  const handleMouseEnter = () => {
    calculatePosition();
    setVisible(true);
  };

  const handleMouseLeave = () => {
    setVisible(false);
  };

  // Close on scroll to avoid detached tooltips
  useEffect(() => {
    if (visible) {
        const handleScroll = () => setVisible(false);
        window.addEventListener('scroll', handleScroll, true);
        return () => window.removeEventListener('scroll', handleScroll, true);
    }
  }, [visible]);

  return (
    <div 
        ref={triggerRef} 
        onMouseEnter={handleMouseEnter} 
        onMouseLeave={handleMouseLeave}
        className={`inline-block ${className}`}
    >
      {children}
      {visible && content && createPortal(
        <div 
            style={{ top: pos.top, left: pos.left }} 
            className={`fixed z-[9999] bg-gray-900 border border-gray-800 text-white text-xs p-3 rounded-lg shadow-xl ${maxWidth} break-words pointer-events-none animate-in fade-in zoom-in-95 duration-100`}
        >
            {content}
        </div>,
        document.body
      )}
    </div>
  );
}
