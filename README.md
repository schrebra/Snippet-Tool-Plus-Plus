
# Snipping Tool++

A powerful, feature-rich PowerShell-based screenshot utility that enhances the Windows screenshot experience with professional-grade features, automatic organization, and customizable borders. Perfect for both casual users and professionals who need a reliable screenshot tool.

![image](https://github.com/user-attachments/assets/cb5ba9cc-3c6a-4b75-80e1-1749339079de)




## Why Choose Snipping Tool++

### For General Users
- **Simpler Than Paint**: No need to paste into Paint and save manually
- **Automatic Organization**: Never lose a screenshot again with automatic sorting
- **Quick Access**: All your screenshots are just one click away
- **Enhanced Visuals**: Add professional-looking borders automatically
- **Zero Configuration**: Works right out of the box with smart defaults
- **Clipboard Ready**: Screenshots are automatically copied to your clipboard
- **User-Friendly**: Simple, intuitive interface that anyone can use

### For Professionals
- **Workflow Integration**: 
  - Timestamp-based naming for easy tracking
  - Consistent file organization
  - Automated border generation for professional documentation
- **Time Saving Features**:
  - No manual editing required for borders
  - Quick access to recent captures
  - Batch processing capabilities
- **Documentation Ready**:
  - Perfect for technical documentation
  - Consistent screenshot styling
  - Professional-grade output
- **Customization Options**:
  - Configurable border properties
  - Adjustable file organization
  - Extensible PowerShell framework

## Features in Detail

### Core Features
- **Precision Capture System**:
  - Click-and-drag selection interface
  - Real-time selection preview
  - High-resolution capture support
  - Multi-monitor compatibility

- **Automatic Organization**:
  - Creates dedicated `Pictures\Snippets` folder
  - Timestamp-based naming (YYMMDD_HHMMSS.png)
  - Automatic file management
  - Easy access to screenshot history

- **Border Generation**:
  - Customizable border colors
  - Adjustable border width (1-10 pixels)
  - Toggle borders on/off
  - Preview border settings

- **User Interface**:
  - Clean, modern design
  - Status notifications
  - Visual feedback
  - Intuitive controls

## How It Works

### Technical Overview
The tool is built using PowerShell and leverages the .NET Framework for its core functionality. It uses:
- System.Windows.Forms for the GUI
- System.Drawing for image manipulation
- Custom C# classes for screen selection
- Native Windows APIs for clipboard operations

### Key Components
1. **Selection Engine**:
   - Transparent overlay for selection
   - Real-time drawing of selection rectangle
   - Cross-hair cursor for precision
   - Right-click cancellation support

2. **Image Processing**:
   - High-quality screenshot capture
   - Border generation algorithm
   - Automatic image formatting
   - Memory-efficient processing

3. **File Management**:
   - Automatic directory creation
   - Unique filename generation
   - File system error handling
   - Path validation

## Usage Guide

### Basic Usage
1. Launch the script
2. Click "New Snippet"
3. Select screen area
4. Screenshot is automatically:
   - Saved to Snippets folder
   - Copied to clipboard
   - Available for immediate viewing

### Advanced Usage
1. **Border Customization**:
   - Enable/disable borders
   - Select border color
   - Adjust border width
   - Preview settings

2. **File Management**:
   - Access recent screenshots
   - Browse snippets folder
   - View capture history
   - Organize captures

## Installation and Setup

### Requirements
- Windows PowerShell 5.1 or higher
- .NET Framework 4.5 or higher
- System.Windows.Forms assembly
- System.Drawing assembly

### Installation Steps
1. Download `Snipping Tool++.ps1`
2. Right-click and select "Run with PowerShell"
3. Optional: Create desktop shortcut
4. Optional: Add to startup

## Best Practices

### For Documentation
- Use consistent border settings
- Establish naming conventions
- Regular backup of Snippets folder
- Organize by project/date

### For Personal Use
- Keep default settings for consistency
- Regular cleanup of old screenshots
- Use keyboard shortcuts when available
- Familiarize with quick access features

## Technical Details

### Architecture
The tool uses a hybrid PowerShell/C# architecture:
- PowerShell for system integration
- C# for performance-critical components
- Windows Forms for UI
- GDI+ for image processing

### Performance
- Optimized for minimal memory usage
- Fast capture and processing
- Efficient file handling
- Responsive user interface

## Future Development

### Planned Features
- Hotkey support
- Additional border styles
- Annotation tools
- Cloud integration
- Multiple output formats
- Batch processing
- Configuration profiles

## Contributing
Contributions are welcome! Feel free to:
- Report bugs
- Suggest features
- Submit pull requests
- Improve documentation

## License
This tool is provided as-is under the MIT license. Feel free to modify and distribute according to your needs.

## Support
For issues, questions, or suggestions:
- Create an issue on GitHub
- Fork and submit improvements
- Share your use cases

## Notes
- Screenshots are saved in PNG format for quality
- The tool creates necessary folders automatically
- Interface remains topmost during capture
- Maintains clipboard contents for easy sharing
