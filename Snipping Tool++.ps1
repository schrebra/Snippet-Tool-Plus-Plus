Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Snippets folder in Pictures if it doesn't exist
$snippetsFolder = Join-Path ([Environment]::GetFolderPath("MyPictures")) "Snippets"
if (-not (Test-Path $snippetsFolder)) {
    New-Item -ItemType Directory -Path $snippetsFolder | Out-Null
}

# Global variable for last screenshot path
$script:lastScreenshotPath = $null

# C# class for transparent form and selection
$csharpCode = @"
using System;
using System.Drawing;
using System.Windows.Forms;

public class SelectionForm : Form
{
    private Point selectStart;
    private Rectangle selection;
    private bool selecting = false;
    public event EventHandler<Rectangle> SelectionComplete;

    public SelectionForm()
    {
        this.TopMost = true;
        this.FormBorderStyle = FormBorderStyle.None;
        this.ShowInTaskbar = false;
        this.StartPosition = FormStartPosition.Manual;
        this.Location = new Point(0, 0);
        this.Size = Screen.PrimaryScreen.Bounds.Size;
        this.BackColor = Color.Black;
        this.Opacity = 0.4;
        this.Cursor = Cursors.Cross;
        this.DoubleBuffered = true;

        this.MouseDown += new MouseEventHandler(SelectionForm_MouseDown);
        this.MouseMove += new MouseEventHandler(SelectionForm_MouseMove);
        this.MouseUp += new MouseEventHandler(SelectionForm_MouseUp);
        this.Paint += new PaintEventHandler(SelectionForm_Paint);
    }

    private void SelectionForm_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            selecting = true;
            selectStart = e.Location;
            selection = new Rectangle();
        }
        else if (e.Button == MouseButtons.Right)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }

    private void SelectionForm_MouseMove(object sender, MouseEventArgs e)
    {
        if (selecting)
        {
            selection = new Rectangle(
                Math.Min(selectStart.X, e.X),
                Math.Min(selectStart.Y, e.Y),
                Math.Abs(e.X - selectStart.X),
                Math.Abs(e.Y - selectStart.Y));

            this.Invalidate();
        }
    }

    private void SelectionForm_MouseUp(object sender, MouseEventArgs e)
    {
        if (selecting && e.Button == MouseButtons.Left)
        {
            selecting = false;
            if (SelectionComplete != null)
            {
                SelectionComplete(this, selection);
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }

    private void SelectionForm_Paint(object sender, PaintEventArgs e)
    {
        if (selecting)
        {
            using (Pen pen = new Pen(Color.Red, 4))
            {
                e.Graphics.DrawRectangle(pen, selection);
            }
        }
    }

    public Rectangle GetSelection()
    {
        return selection;
    }
}
"@

# Compile the C# code
Add-Type -TypeDefinition $csharpCode -ReferencedAssemblies System.Windows.Forms, System.Drawing

# Function to capture screenshot
function Capture-Screenshot {
    param (
        [System.Drawing.Rectangle]$bounds
    )

    $screenshot = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
    $graphics = [System.Drawing.Graphics]::FromImage($screenshot)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
    $graphics.Dispose()
    return $screenshot
}

# Function to copy image to clipboard
function Copy-ImageToClipboard {
    param (
        [System.Drawing.Image]$image
    )

    if ($image -ne $null) {
        [System.Windows.Forms.Clipboard]::SetImage($image)
    }
}

# Function to add outline to image
function Add-ImageOutline {
    param (
        [System.Drawing.Image]$image,
        [System.Drawing.Color]$color,
        [int]$width
    )

    $newWidth = $image.Width + ($width * 2)
    $newHeight = $image.Height + ($width * 2)
    $newImage = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
    $graphics = [System.Drawing.Graphics]::FromImage($newImage)

    # Fill background with outline color
    $graphics.Clear($color)

    # Draw original image in center
    $graphics.DrawImage($image, $width, $width, $image.Width, $image.Height)

    $graphics.Dispose()
    return $newImage
}
# Custom button style function
function Set-ModernButtonStyle {
    param($button)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 1
    $button.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#DEE2E6")
    $button.BackColor = [System.Drawing.Color]::White
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    # Hover effect
    $button.Add_MouseEnter({
        $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F1F3F5")
    })
    $button.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::White
    })
}

# Create the main form with modern styling
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = " Snipping Tool++"
$mainForm.Size = New-Object System.Drawing.Size(585, 280)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.MaximizeBox = $false
$mainForm.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F8F9FA")
$mainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Create a container panel for better organization
$containerPanel = New-Object System.Windows.Forms.Panel
$containerPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$containerPanel.Padding = New-Object System.Windows.Forms.Padding(20)
$mainForm.Controls.Add($containerPanel)

# Button dimensions and spacing
$buttonWidth = 160
$buttonHeight = 45
$buttonSpacing = 20
$totalButtonsWidth = ($buttonWidth * 3) + ($buttonSpacing * 2)

# Calculate positions for centered buttons
$firstButtonX = ($mainForm.ClientSize.Width - $totalButtonsWidth + 10) / 2
$secondButtonX = $firstButtonX + $buttonWidth + $buttonSpacing
$thirdButtonX = $secondButtonX + $buttonWidth + $buttonSpacing
$buttonY = 20

# Create the New Snip button with modern styling
$newSnipButton = New-Object System.Windows.Forms.Button
$newSnipButton.Location = New-Object System.Drawing.Point([int]$firstButtonX, $buttonY)
$newSnipButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
$newSnipButton.Text = "New Snippet"
Set-ModernButtonStyle $newSnipButton
$newSnipButton.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml("#228BE6")
$newSnipButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#228BE6")
$containerPanel.Controls.Add($newSnipButton)

# Create the Open with Photo Viewer button
$openButton = New-Object System.Windows.Forms.Button
$openButton.Location = New-Object System.Drawing.Point([int]$secondButtonX, $buttonY)
$openButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
$openButton.Text = "Open Last Snippet"
$openButton.Enabled = $false
Set-ModernButtonStyle $openButton
$containerPanel.Controls.Add($openButton)

# Create the Open Snippets Folder button
$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Location = New-Object System.Drawing.Point([int]$thirdButtonX, $buttonY)
$openFolderButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
$openFolderButton.Text = "Open Snippets Folder"
Set-ModernButtonStyle $openFolderButton
$containerPanel.Controls.Add($openFolderButton)

# Create modern outline options group
$outlineGroup = New-Object System.Windows.Forms.GroupBox
$outlineGroup.Text = "Outline Options"
$outlineGroup.Location = New-Object System.Drawing.Point(20, ($buttonY + $buttonHeight + 25))
$outlineGroup.Size = New-Object System.Drawing.Size(540, 100)
$outlineGroup.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$outlineGroup.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F8F9FA")

$containerPanel.Controls.Add($outlineGroup)

# Calculate center positions for controls
$totalControlsWidth = 400  # Total width of all controls combined
$startX = ($outlineGroup.Width - $totalControlsWidth) / 2
$centerY = ($outlineGroup.Height - 15) / 2  # 30 is approximate height of controls

# Create outline checkbox with modern styling
$outlineCheckbox = New-Object System.Windows.Forms.CheckBox
$outlineCheckbox.Text = "Add Outline"
$outlineCheckbox.Location = New-Object System.Drawing.Point($startX, $centerY)
$outlineCheckbox.Size = New-Object System.Drawing.Size(100, 24)
$outlineCheckbox.Checked = $true
$outlineGroup.Controls.Add($outlineCheckbox)

# Create modern color button
$colorButton = New-Object System.Windows.Forms.Button
$colorButton.Text = "Select Color"
$colorButton.Location = New-Object System.Drawing.Point(($startX + 120), $centerY)
$colorButton.Size = New-Object System.Drawing.Size(90, 28)
Set-ModernButtonStyle $colorButton
$outlineGroup.Controls.Add($colorButton)

# Create color preview panel with modern styling
$colorPreview = New-Object System.Windows.Forms.Panel
$colorPreview.Location = New-Object System.Drawing.Point(($startX + 220), $centerY)
$colorPreview.Size = New-Object System.Drawing.Size(28, 28)
$colorPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$colorPreview.BackColor = [System.Drawing.Color]::Black
$outlineGroup.Controls.Add($colorPreview)

# Create width label with modern styling
$widthLabel = New-Object System.Windows.Forms.Label
$widthLabel.Text = "Width:"
$widthLabel.Location = New-Object System.Drawing.Point(($startX + 280), ($centerY + 4))
$widthLabel.Size = New-Object System.Drawing.Size(45, 20)
$outlineGroup.Controls.Add($widthLabel)

# Create modern width numeric updown
$widthUpDown = New-Object System.Windows.Forms.NumericUpDown
$widthUpDown.Location = New-Object System.Drawing.Point(($startX + 330), $centerY)
$widthUpDown.Size = New-Object System.Drawing.Size(60, 28)
$widthUpDown.Minimum = 1
$widthUpDown.Maximum = 10
$widthUpDown.Value = 2
$widthUpDown.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$outlineGroup.Controls.Add($widthUpDown)

# Create modern status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, ($outlineGroup.Location.Y + $outlineGroup.Height + 15))
$statusLabel.Size = New-Object System.Drawing.Size(540, 30)
$statusLabel.Text = "Click 'New Snippet' to start. Right-click to cancel."
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$statusLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E9ECEF")
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$containerPanel.Controls.Add($statusLabel)

# Event handlers
$newSnipButton.Add_Click({
    $mainForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
    Start-Sleep -Milliseconds 250

    $selectionForm = New-Object SelectionForm
    $result = $selectionForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selection = $selectionForm.GetSelection()
        
        if ($selection.Width -gt 0 -and $selection.Height -gt 0) {
            $screenshot = Capture-Screenshot -bounds $selection
            
            if ($outlineCheckbox.Checked) {
                $screenshot = Add-ImageOutline -image $screenshot -color $colorPreview.BackColor -width $widthUpDown.Value
            }
            
            # Generate filename with timestamp
            $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
            $filename = "Screenshot_$timestamp.png"
            $filepath = Join-Path $snippetsFolder $filename
            
            # Save the image
            $screenshot.Save($filepath, [System.Drawing.Imaging.ImageFormat]::Png)
            
            # Copy to clipboard
            Copy-ImageToClipboard -image $screenshot
            
            # Update status and enable open button
            $statusLabel.Text = "Screenshot saved as $filename and copied to clipboard"
            $openButton.Enabled = $true
            $script:lastScreenshotPath = $filepath
            
            # Cleanup
            $screenshot.Dispose()
        }
    }
    
    $mainForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
})

$openButton.Add_Click({
    if ($script:lastScreenshotPath -and (Test-Path $script:lastScreenshotPath)) {
        Start-Process $script:lastScreenshotPath
    }
})

$openFolderButton.Add_Click({
    Start-Process $snippetsFolder
})

$colorButton.Add_Click({
    $colorDialog = New-Object System.Windows.Forms.ColorDialog
    $colorDialog.Color = $colorPreview.BackColor
    
    if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $colorPreview.BackColor = $colorDialog.Color
    }
})

# Clean up resources when form closes
$mainForm.Add_FormClosing({
    if ($script:lastScreenshotPath) {
        $screenshot.Dispose()
    }
})

# Show the form
$mainForm.Add_Shown({$mainForm.Activate()})
[void]$mainForm.ShowDialog()

# Clean up when script ends
$mainForm.Dispose()