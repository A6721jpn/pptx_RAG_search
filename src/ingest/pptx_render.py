import win32com.client
from pathlib import Path
from typing import List
import logging
import os

logger = logging.getLogger(__name__)

class PptxRenderer:
    """Renders PowerPoint slides to images using Windows COM."""
    
    def render(self, file_path: Path, output_dir: Path) -> List[str]:
        """
        Render all slides in the PPTX to PNG files in the output directory.
        
        Args:
            file_path: Absolute path to the PPTX file
            output_dir: Absolute path to the output directory
            
        Returns:
            List of absolute paths to the generated image files.
        """
        file_path = file_path.resolve()
        output_dir = output_dir.resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            # Dispatch PowerPoint application
            # Use DispatchEx to possibly get a fresh instance? Dispatch is usually fine.
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            
            # Opening WithWindow=False often helps with background processing, 
            # but some PPT versions require it to be True or will fail to export.
            # We try WithWindow=False (msoFalse = 0)
            presentation = powerpoint.Presentations.Open(str(file_path), WithWindow=0, ReadOnly=True)
            
            image_paths = []
            for i, slide in enumerate(presentation.Slides):
                slide_no = i + 1
                image_name = f"slide_{slide_no:04d}.png"
                image_path = output_dir / image_name
                
                # Export to PNG
                # 3rd and 4th args are width/height. 0=default (slide size)
                slide.Export(str(image_path), "PNG", 0, 0)
                image_paths.append(str(image_path))
            
            presentation.Close()
            return image_paths
            
        except Exception as e:
            logger.error(f"COM render failed for {file_path}: {e}")
            # Try to release if possible, but mainly rely on finally block in caller or persistent app
            raise
