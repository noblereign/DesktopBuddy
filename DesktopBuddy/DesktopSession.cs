using System;
using System.Collections.Generic;
using FrooxEngine;
using FrooxEngine.UIX;

namespace DesktopBuddy;

public class DesktopSession
{
    public DesktopStreamer Streamer;
    public SolidColorTexture Texture;
    public Canvas Canvas;
    public Slot Root;
    public bool UpdateInProgress;
    public bool ManualModeSet;
    public double TimeSinceLastCapture;
    public double TargetInterval;

    // VR hand tracking: which source last clicked/scrolled — only this source moves the mouse
    public Component LastActiveSource;

    // Scroll: direction tracking to suppress jitter, tick tracking to prevent burst
    public int LastScrollSign;
    public double LastScrollTick;

    // One-shot diagnostic flag for joystick detection
    public bool JoystickDiagLogged;

    // Unique stream ID (never reused, never shifts)
    public int StreamId;
    public IntPtr Hwnd; // For shared stream cleanup

    // Downscaled frame buffer for stream encoding
    public byte[] ScaledBuffer;

    // Stream: shared frame for FFmpeg encoding (set by update loop, read by encoder thread)
    public volatile byte[] StreamFrame;
    public int StreamWidth, StreamHeight;
    public readonly object StreamLock = new();
    public volatile bool StreamFrameReady;

    // Child window tracking — popups, dialogs, context menus from the same process
    public uint ProcessId;
    public double TimeSinceChildCheck;
    public HashSet<IntPtr> TrackedChildHwnds = new();
    public List<DesktopSession> ChildSessions = new();
    public DesktopSession ParentSession; // non-null if this IS a child panel
    public bool IsChildPanel => ParentSession != null;
    public bool Cleaned; // Guard: CleanupSession already ran for this session
}
