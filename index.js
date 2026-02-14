/**
 * Manual ToolChanger
 * Manual tool change workflow with Tool Length Setter (TLS) support.
 *
 * Copyright (C) 2024 Francis Marasigan
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

const toFiniteNumber = (value, fallback = 0) => {
  const num = Number.parseFloat(value);
  return Number.isFinite(num) ? num : fallback;
};

const sanitizeCoords = (coords = {}) => ({
  x: toFiniteNumber(coords.x),
  y: toFiniteNumber(coords.y),
  z: toFiniteNumber(coords.z)
});

const buildInitialConfig = (raw = {}) => ({
  // Position Settings
  pocket1: sanitizeCoords(raw.pocket1),
  toolSetter: sanitizeCoords(raw.toolSetter),
  parking: sanitizeCoords(raw.parking),

  // Tool Settings
  numberOfTools: toFiniteNumber(raw.numberOfTools, 1),

  // UI Toggle Settings
  autoSwap: raw.autoSwap === true,
  pauseBeforeUnload: raw.pauseBeforeUnload !== false,
  showMacroCommand: raw.showMacroCommand ?? false,
  performTlsAfterHome: raw.performTlsAfterHome ?? false,
  waitForSpindle: raw.waitForSpindle !== false,

  // Advanced Settings (no UI, JSON only)
  // Z-axis Settings
  zEngagement: toFiniteNumber(raw.zEngagement, -50),
  zSafe: toFiniteNumber(raw.zSafe, 0),
  zSpinOff: toFiniteNumber(raw.zSpinOff, 23),
  zRetreat: toFiniteNumber(raw.zRetreat, 12),

  // Tool Change Settings
  unloadRpm: toFiniteNumber(raw.unloadRpm, 1500),
  loadRpm: toFiniteNumber(raw.loadRpm, 1200),
  engageFeedrate: toFiniteNumber(raw.engageFeedrate, 3500),

  // Tool Length Setter Settings
  zProbeStart: toFiniteNumber(raw.zProbeStart, -10),
  seekDistance: toFiniteNumber(raw.seekDistance, 50),
  seekFeedrate: toFiniteNumber(raw.seekFeedrate, 100),

  // Aux Output Settings (-1 = disabled, 0+ = M64 P{n}, 'M7' or 'M8' for coolant)
  tlsAuxOutput: raw.tlsAuxOutput === 'M7' || raw.tlsAuxOutput === 'M8'
    ? raw.tlsAuxOutput
    : toFiniteNumber(raw.tlsAuxOutput, -1),

  // Tool Change Events
  preToolChangeGcode: raw.preToolChangeGcode ?? '',
  postToolChangeGcode: raw.postToolChangeGcode ?? '',
  abortEventGcode: raw.abortEventGcode ?? ''
});

const resolveServerPort = (pluginSettings = {}, appSettings = {}) => {
  const appPort = Number.parseInt(appSettings?.senderPort, 10);
  if (Number.isFinite(appPort)) {
    return appPort;
  }

  const pluginPort = Number.parseInt(pluginSettings?.port, 10);
  if (Number.isFinite(pluginPort)) {
    return pluginPort;
  }

  return 8090;
};

// Helper: Fetch tool offsets from tool library
async function getToolOffsets(toolNumber, ctx) {
  if (!toolNumber || toolNumber <= 0) {
    return { x: 0, y: 0, z: 0 };
  }
  try {
    const pluginSettings = ctx.getSettings() || {};
    const appSettings = ctx.getAppSettings() || {};
    const port = resolveServerPort(pluginSettings, appSettings);
    const response = await fetch(`http://localhost:${port}/api/tools`);
    if (response.ok) {
      const tools = await response.json();
      const tool = tools.find(t => t.toolNumber === toolNumber);
      if (tool && tool.offsets) {
        return { x: tool.offsets.x || 0, y: tool.offsets.y || 0, z: tool.offsets.z || 0 };
      }
    }
  } catch (error) {
    ctx.log('Failed to fetch tool offsets:', error);
  }
  return { x: 0, y: 0, z: 0 };
}

// Helper: Format and split G-code into array of commands
const formatGCode = (gcode) => {
  return gcode
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.length > 0);
};

// Helper: Create tool length setter routine
function createToolLengthSetRoutine(settings, toolOffsets = { x: 0, y: 0, z: 0 }) {
  const tlsX = settings.toolSetter.x + (toolOffsets.x || 0);
  const tlsY = settings.toolSetter.y + (toolOffsets.y || 0);
  const tlsZ = toolOffsets.z || 0;
  const fineProbeFeedrate = settings.seekFeedrate < 75 ? settings.seekFeedrate : 75;

  // Extra Z rapid move for shorter tools (z offset is typically negative)
  const extraZMove = tlsZ !== 0 ? `G91 G0 Z${tlsZ}\n    G90` : '';

  // Aux output switching during TLS (if configured)
  const auxOutput = settings.tlsAuxOutput;
  let auxOn = '';
  let auxOff = '';
  if (auxOutput === 'M7' || auxOutput === 'M8') {
    auxOn = `G4 P0\n    ${auxOutput}\n    G4 P0`;
    auxOff = `G4 P0\n    M9\n    G4 P0`;
  } else if (typeof auxOutput === 'number' && auxOutput >= 0) {
    auxOn = `G4 P0\n    M64 P${auxOutput}\n    G4 P0`;
    auxOff = `G4 P0\n    M65 P${auxOutput}\n    G4 P0`;
  }

  return `
    G53 G0 Z${settings.zSafe}
    G53 G0 X${tlsX} Y${tlsY}
    G53 G0 Z${settings.toolSetter.z}
    ${extraZMove}
    ${auxOn}
    G43.1 Z0
    G38.2 G91 Z-${settings.seekDistance} F${settings.seekFeedrate}
    G38.4 G91 Z5 F${fineProbeFeedrate}
    G91 G0 Z5
    G90
    ${auxOff}
    #<_ofs_idx> = [#5220 * 20 + 5203]
    #<_cur_wcs_z_ofs> = #[#<_ofs_idx>]
    #<_rc_trigger_mach_z> = [#5063 + #<_cur_wcs_z_ofs>]
    G43.1 Z[#<_rc_trigger_mach_z>]
    (Notify ncSender that toolLengthSet is now set)
    $#=_tool_offset
    G53 G91 G0 Z${settings.zSafe}
  `.trim();
}

// Helper: Check if a tool number is a manual tool (no RCS pocket)
function isManualTool(toolNumber, settings) {
  return toolNumber > settings.numberOfTools;
}

// Helper: Tool unload routine
function createToolUnload(settings, currentTool, targetTool) {
  const useRCS = settings.autoSwap && !isManualTool(currentTool, settings);
  const messageCode = useRCS ? 'PLUGIN_RCS:UNLOAD_MESSAGE' : 'PLUGIN_RCS:UNLOAD_MESSAGE_MANUAL';
  const needsPause = settings.pauseBeforeUnload;
  const confirmationLines = needsPause ? `
    G4 P0
    (MSG, ${messageCode})
    M0` : '';

  if (useRCS) {
    // RapidChangeSolo ON
    const pauseSequence = needsPause ? `
      G53 G0 X${settings.parking.x} Y${settings.parking.y}
      ${confirmationLines}
      G53 G0 X${settings.pocket1.x} Y${settings.pocket1.y}` : `
      G53 G0 X${settings.pocket1.x} Y${settings.pocket1.y}`;

    // Only move to manual location after unload if we're swapping tools (targetTool !== 0)
    const moveToManualAfterUnload = targetTool !== 0 ? `
      G53 G0 X${settings.parking.x} Y${settings.parking.y}` : '';

    // Insert G65P6 to skip spindle wait when waitForSpindle is disabled
    const g65p6 = settings.waitForSpindle ? '' : 'G65P6';

    return `
      G53 G0 Z${settings.zSafe}
      ${pauseSequence}
      G53 G0 Z${settings.zEngagement + settings.zSpinOff}
      ${g65p6}
      M4 S${settings.unloadRpm}
      G53 G1 Z${settings.zEngagement} F${settings.engageFeedrate}
      G53 G1 Z${settings.zEngagement + settings.zRetreat} F${settings.engageFeedrate}
      ${g65p6}
      M5
      M61 Q0
      G53 G0 Z${settings.zSafe}
      ${moveToManualAfterUnload}
    `.trim();
  } else {
    // RapidChangeSolo OFF: Move to manual tool change location with Z0
    // Only show confirmation if unloading only (not swapping) - swap shows combined message in load
    const isUnloadOnly = targetTool === 0;
    const manualConfirmation = isUnloadOnly ? `
      G4 P0
      (MSG, ${messageCode})
      M0` : '';

    return `
      G53 G0 Z${settings.zSafe}
      G53 G0 X${settings.parking.x} Y${settings.parking.y}
      G53 G0 Z0
      ${manualConfirmation}
      M61 Q0
    `.trim();
  }
}

// Helper: Tool load routine
function createToolLoad(settings, toolNumber, hasUnload, currentTool) {
  const useRCS = settings.autoSwap && !isManualTool(toolNumber, settings);
  // Only use swap message when the unload was also manual (user had to remove the bit by hand)
  // If unload was via RCS, the bit is already removed automatically
  const wasManualUnload = hasUnload && isManualTool(currentTool, settings);
  const messageCode = useRCS
    ? `PLUGIN_RCS:LOAD_MESSAGE_${toolNumber}`
    : (wasManualUnload ? `PLUGIN_RCS:SWAP_MESSAGE_MANUAL_${toolNumber}` : `PLUGIN_RCS:LOAD_MESSAGE_MANUAL_${toolNumber}`);

  // If no unload happened, we need to move to manual location first
  const moveToManualLocation = hasUnload ? '' : `
      G53 G0 Z${settings.zSafe}
      G53 G0 X${settings.parking.x} Y${settings.parking.y}`;

  if (useRCS) {
    // RapidChangeSolo ON: Go to manual location, show message, then move to pocket1 to load
    // Insert G65P6 to skip spindle wait when waitForSpindle is disabled
    const g65p6 = settings.waitForSpindle ? '' : 'G65P6';

    return `
      ${moveToManualLocation}
      G4 P0
      (MSG, ${messageCode})
      M0
      G53 G0 Z${settings.zSafe}
      G53 G0 X${settings.pocket1.x} Y${settings.pocket1.y}
      G53 G0 Z${settings.zEngagement + settings.zSpinOff}
      ${g65p6}
      M3 S${settings.loadRpm}
      G53 G1 Z${settings.zEngagement} F${settings.engageFeedrate}
      G53 G1 Z${settings.zEngagement + settings.zRetreat} F${settings.engageFeedrate}
      G53 G1 Z${settings.zEngagement} F${settings.engageFeedrate}
      G53 G1 Z${settings.zEngagement + settings.zRetreat} F${settings.engageFeedrate}
      G53 G1 Z${settings.zEngagement} F${settings.engageFeedrate}
      G53 G1 Z${settings.zEngagement + settings.zRetreat} F${settings.engageFeedrate}
      ${g65p6}
      M5
      M61 Q${toolNumber}
      G53 G0 Z${settings.zSafe}
    `.trim();
  } else {
    // RapidChangeSolo OFF: Go to manual location with Z0, show message, user installs tool
    return `
      ${moveToManualLocation}
      G53 G0 Z0
      G4 P0
      (MSG, ${messageCode})
      M0
      M61 Q${toolNumber}
      G53 G0 Z${settings.zSafe}
    `.trim();
  }
}

// Build unload tool section
function buildUnloadTool(settings, currentTool, targetTool) {
  if (currentTool === 0) {
    return '';
  }
  return `
    (Unload current tool T${currentTool})
    ${createToolUnload(settings, currentTool, targetTool)}
  `.trim();
}

// Build load tool section
function buildLoadTool(settings, toolNumber, tlsRoutine, hasUnload, currentTool) {
  if (toolNumber === 0) {
    return '';
  }
  return `
    (Load new tool T${toolNumber})
    ${createToolLoad(settings, toolNumber, hasUnload, currentTool)}
    ${tlsRoutine}
  `.trim();
}

// Helper: Build tool change program
function buildToolChangeProgram(settings, currentTool, toolNumber, toolOffsets = { x: 0, y: 0 }) {
  const tlsRoutine = createToolLengthSetRoutine(settings, toolOffsets);
  const hasUnload = currentTool !== 0;

  // Build sections
  const unloadSection = buildUnloadTool(settings, currentTool, toolNumber);
  const loadSection = buildLoadTool(settings, toolNumber, tlsRoutine, hasUnload, currentTool);

  // Tool change event commands
  const preToolChangeCmd = settings.preToolChangeGcode?.trim() || '';
  const postToolChangeCmd = settings.postToolChangeGcode?.trim() || '';

  // Assemble complete program with return to units wrapper
  const gcode = `
    (Start of Manual ToolChanger Sequence)
    ${preToolChangeCmd}
    #<return_units> = [20 + #<_metric>]
    G21
    M5
    ${unloadSection}
    ${loadSection}
    G53 G0 Z${settings.zSafe}
    G[#<return_units>]
    G90
    ${postToolChangeCmd}
    (End of Manual ToolChanger Sequence)
  `.trim();

  return formatGCode(gcode);
}

// Handle $TLS command
async function handleTLSCommand(commands, context, settings, ctx) {
  const tlsIndex = commands.findIndex(cmd =>
    cmd.isOriginal && cmd.command.trim().toUpperCase() === '$TLS'
  );

  if (tlsIndex === -1) {
    return; // No $TLS command found
  }

  ctx.log('$TLS command detected, replacing with tool length setter routine');

  // Get current tool and fetch its TLS offsets
  const currentTool = context.machineState?.tool ?? 0;
  const toolOffsets = await getToolOffsets(currentTool, ctx);
  if (toolOffsets.x !== 0 || toolOffsets.y !== 0 || toolOffsets.z !== 0) {
    ctx.log(`Applying TLS offsets for T${currentTool}: X=${toolOffsets.x}, Y=${toolOffsets.y}, Z=${toolOffsets.z}`);
  }

  const tlsCommand = commands[tlsIndex];
  const toolLengthSetRoutine = createToolLengthSetRoutine(settings, toolOffsets);

  // Pre/Post event commands (same as M6 tool change)
  const preToolChangeCmd = settings.preToolChangeGcode?.trim() || '';
  const postToolChangeCmd = settings.postToolChangeGcode?.trim() || '';

  const gcode = `
    (Start of Tool Length Setter)
    ${preToolChangeCmd}
    #<return_units> = [20 + #<_metric>]
    G21
    ${toolLengthSetRoutine}
    G53 G0 Z${settings.zSafe}
    G[#<return_units>]
    G90
    ${postToolChangeCmd}
    (End of Tool Length Setter)
  `.trim();
  const tlsProgram = formatGCode(gcode);
  const showMacroCommand = settings.showMacroCommand ?? false;

  const expandedCommands = tlsProgram.map((line, index) => {
    if (index === 0) {
      // First command - show $TLS if hiding macro, otherwise show actual command
      return {
        command: line,
        displayCommand: showMacroCommand ? null : tlsCommand.command.trim(),
        isOriginal: false
      };
    } else {
      // Rest of commands - hide if not showing macro
      return {
        command: line,
        displayCommand: null,
        isOriginal: false,
        meta: showMacroCommand ? {} : { silent: true }
      };
    }
  });

  commands.splice(tlsIndex, 1, ...expandedCommands);
}

// Handle $POCKET1 command
function handlePocket1Command(commands, settings, ctx) {
  const pocket1Index = commands.findIndex(cmd =>
    cmd.isOriginal && cmd.command.trim().toUpperCase() === '$POCKET1'
  );

  if (pocket1Index === -1) {
    return; // No $POCKET1 command found
  }

  ctx.log('$POCKET1 command detected, moving to pocket location');

  const pocket1Command = commands[pocket1Index];
  const gcode = `
    G53 G21 G90 G0 Z${settings.zSafe}
    G53 G21 G90 G0 X${settings.pocket1.x} Y${settings.pocket1.y}
  `.trim();

  const pocket1Program = formatGCode(gcode);
  const showMacroCommand = settings.showMacroCommand ?? false;

  const expandedCommands = pocket1Program.map((line, index) => {
    if (index === 0) {
      // First command - show $POCKET1 if hiding macro, otherwise show actual command
      return {
        command: line,
        displayCommand: showMacroCommand ? null : pocket1Command.command.trim(),
        isOriginal: false
      };
    } else {
      // Rest of commands - hide if not showing macro
      return {
        command: line,
        displayCommand: null,
        isOriginal: false,
        meta: showMacroCommand ? {} : { silent: true }
      };
    }
  });

  commands.splice(pocket1Index, 1, ...expandedCommands);
}

// Handle $H (home) command with conditional TLS
async function handleHomeCommand(commands, context, settings, ctx) {
  const homeIndex = commands.findIndex(cmd =>
    cmd.isOriginal && cmd.command.trim().toUpperCase() === '$H'
  );

  if (homeIndex === -1) {
    return; // No $H command found
  }

  // Only handle if performTlsAfterHome is enabled
  if (!settings.performTlsAfterHome) {
    return;
  }

  ctx.log('$H command detected with performTlsAfterHome enabled, adding conditional TLS');

  // Get current tool and fetch its TLS offsets
  const currentTool = context.machineState?.tool ?? 0;
  const toolOffsets = await getToolOffsets(currentTool, ctx);
  if (toolOffsets.x !== 0 || toolOffsets.y !== 0 || toolOffsets.z !== 0) {
    ctx.log(`Applying TLS offsets for T${currentTool}: X=${toolOffsets.x}, Y=${toolOffsets.y}, Z=${toolOffsets.z}`);
  }

  const homeCommand = commands[homeIndex];
  const tlsRoutine = createToolLengthSetRoutine(settings, toolOffsets);

  // Pre/Post event commands (only run if TLS runs)
  const preToolChangeCmd = settings.preToolChangeGcode?.trim() || '';
  const postToolChangeCmd = settings.postToolChangeGcode?.trim() || '';

  const gcode = `
    $H
    #<return_units> = [20 + #<_metric>]
    o100 IF [[#<_tool_offset> EQ 0] AND [#<_current_tool> NE 0]]
      ${preToolChangeCmd}
      G21
      ${tlsRoutine}
      G53 G0 Z${settings.zSafe}
      G4 P0
      G53 G0 X0 Y0
      ${postToolChangeCmd}
    o100 ENDIF
    G[#<return_units>]
  `.trim();

  const homeProgram = formatGCode(gcode);
  const showMacroCommand = settings.showMacroCommand ?? false;

  const expandedCommands = homeProgram.map((line, index) => {
    if (index === 0) {
      return {
        command: line,
        displayCommand: showMacroCommand ? null : homeCommand.command.trim(),
        isOriginal: false
      };
    } else {
      return {
        command: line,
        displayCommand: null,
        isOriginal: false,
        meta: showMacroCommand ? {} : { silent: true }
      };
    }
  });

  commands.splice(homeIndex, 1, ...expandedCommands);
}

// Show safety warning dialog
function showSafetyWarningDialog(ctx, title, message, continueLabel, abortEventGcode = '') {
  const abortGcodeLines = abortEventGcode ? abortEventGcode.trim().split('\\n').filter(line => line.trim()) : [];
  const abortGcodeJson = JSON.stringify(abortGcodeLines);
  ctx.showModal(
    /* html */ `
      <style>
        .rcs-safety-container {
          background: var(--color-surface);
          border-radius: var(--radius-medium);
          padding: 32px;
          max-width: 500px;
          box-shadow: 0 20px 60px rgba(0, 0, 0, 0.5);
        }

        .rcs-safety-header {
          font-size: 1.5rem;
          font-weight: 600;
          color: var(--color-text-primary);
          margin-bottom: 24px;
          text-align: center;
        }

        .rcs-safety-dialog {
          display: flex;
          flex-direction: column;
          gap: 24px;
        }

        .rcs-safety-message {
          font-size: 1rem;
          line-height: 1.5;
          color: var(--color-text-primary);
          background: color-mix(in srgb, var(--color-warning) 15%, transparent);
          border: 2px solid var(--color-warning);
          border-radius: var(--radius-small);
          padding: 16px;
        }

        .rcs-safety-actions {
          display: flex;
          justify-content: center;
          gap: 16px;
        }

        .rcs-action-button {
          padding: 12px 32px;
          border: none;
          border-radius: var(--radius-small);
          font-weight: 600;
          font-size: 1rem;
          cursor: pointer;
          transition: all 0.2s ease;
          min-width: 140px;
        }

        .rcs-action-button:hover {
          opacity: 0.9;
        }

        .rcs-action-button:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .rcs-button-abort {
          background: var(--color-error, #dc2626);
          color: white;
        }

        .rcs-button-continue {
          background: var(--color-success, #16a34a);
          color: white;
        }
      </style>

      <div class="rcs-safety-container">
        <div class="rcs-safety-header">${title}</div>
        <div class="rcs-safety-dialog">
          <div class="rcs-safety-message">${message}</div>
          <div class="rcs-safety-actions">
            <button class="rcs-action-button rcs-button-abort" id="rcs-abort-btn">Abort</button>
            <button class="rcs-action-button rcs-button-continue" id="rcs-continue-btn">${continueLabel}</button>
          </div>
        </div>
      </div>

      <script>
        (function() {
          const abortBtn = document.getElementById('rcs-abort-btn');
          const continueBtn = document.getElementById('rcs-continue-btn');
          const abortGcodeLines = ${abortGcodeJson};

          abortBtn.addEventListener('click', function() {
            if (abortBtn.disabled) return;
            abortBtn.disabled = true;
            continueBtn.disabled = true;

            // Execute abort event G-code first if configured
            if (abortGcodeLines.length > 0) {
              abortGcodeLines.forEach(function(line) {
                window.postMessage({
                  type: 'send-command',
                  command: line,
                  displayCommand: line
                }, '*');
              });
            }

            // Then send soft reset to stop everything
            window.postMessage({
              type: 'send-command',
              command: '\\x18',
              displayCommand: '\\x18 (Soft Reset)'
            }, '*');

            window.postMessage({
              type: 'send-command',
              command: '$NCSENDER_CLEAR_MSG',
              displayCommand: '$NCSENDER_CLEAR_MSG'
            }, '*');
          });

          continueBtn.addEventListener('click', function() {
            if (continueBtn.disabled) return;
            abortBtn.disabled = true;
            continueBtn.disabled = true;

            window.postMessage({
              type: 'send-command',
              command: '~',
              displayCommand: '~ (Cycle Start)'
            }, '*');

            window.postMessage({
              type: 'send-command',
              command: '$NCSENDER_CLEAR_MSG',
              displayCommand: '$NCSENDER_CLEAR_MSG'
            }, '*');
          });
        })();
      </script>
    `,
    { closable: false }
  );
}

// Handle M6 tool change command
async function handleM6Command(commands, context, settings, ctx) {
  // Find original M6 command
  const m6Index = commands.findIndex(cmd => {
    if (!cmd.isOriginal) return false;
    const parsed = ctx.utils.parseM6Command(cmd.command);
    return parsed?.matched && parsed.toolNumber !== null;
  });

  if (m6Index === -1) {
    return; // No M6 found
  }

  const m6Command = commands[m6Index];
  const parsed = ctx.utils.parseM6Command(m6Command.command);

  if (!parsed?.matched || parsed.toolNumber === null) {
    return;
  }

  const toolNumber = parsed.toolNumber;
  const location = context.lineNumber !== undefined ? `at line ${context.lineNumber}` : `from ${context.sourceId}`;
  const currentTool = context.machineState?.tool ?? 0;

  // Fetch TLS offsets for the NEW tool being loaded
  const toolOffsets = await getToolOffsets(toolNumber, ctx);
  if (toolOffsets.x !== 0 || toolOffsets.y !== 0 || toolOffsets.z !== 0) {
    ctx.log(`Applying TLS offsets for T${toolNumber}: X=${toolOffsets.x}, Y=${toolOffsets.y}, Z=${toolOffsets.z}`);
  }

  ctx.log(`M6 detected with tool T${toolNumber} ${location}, current tool: T${currentTool}, executing tool change program`);

  const toolChangeProgram = buildToolChangeProgram(settings, currentTool, toolNumber, toolOffsets);
  const showMacroCommand = settings.showMacroCommand ?? false;

  // Replace M6 command with expanded program
  const expandedCommands = toolChangeProgram.map((line, index) => {
    if (index === 0) {
      // First command - show M6 if hiding macro, otherwise show actual command
      return {
        command: line,
        displayCommand: showMacroCommand ? null : m6Command.command.trim(),
        isOriginal: false
      };
    } else {
      // Rest of commands - hide if not showing macro
      return {
        command: line,
        displayCommand: null,
        isOriginal: false,
        meta: showMacroCommand ? {} : { silent: true }
      };
    }
  });

  commands.splice(m6Index, 1, ...expandedCommands);
}

export async function onLoad(ctx) {
  ctx.log('Manual ToolChanger loaded');

  // Get current plugin settings and app settings
  const pluginSettings = ctx.getSettings() || {};
  const appSettings = ctx.getAppSettings() || {};

  // Check if plugin has been configured (has required settings)
  const isConfigured = !!(pluginSettings.pocket1);

  // Get saved number of tools (default to 1 if not set)
  const toolCount = pluginSettings.numberOfTools || 1;

  // Set tool.source to indicate this plugin controls the tool settings
  // Only enable manual and TLS tools if plugin is configured
  await new Promise(resolve => setTimeout(resolve, 100));

  try {
    const response = await fetch(`http://localhost:${resolveServerPort(pluginSettings, appSettings)}/api/settings`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tool: {
          count: toolCount,
          source: 'com.ncsender.manualtoolchange',
          manual: isConfigured,
          tls: isConfigured
        }
      })
    });

    if (response.ok) {
      ctx.log(`Tool settings synchronized: count=${toolCount}, manual=${isConfigured}, tls=${isConfigured} (source: com.ncsender.manualtoolchange)`);
    } else {
      ctx.log(`Failed to sync tool settings: ${response.status}`);
    }
  } catch (error) {
    ctx.log('Failed to sync tool settings on plugin load:', error);
  }

  const MESSAGE_MAP = {
    'PLUGIN_RCS:LOAD_MESSAGE_MANUAL': {
      title: 'Loading',
      message: 'Please install {toolNumber} securely, then click <em>"Continue"</em> to proceed or <em>"Abort"</em> to cancel.',
      continueLabel: 'Continue'
    },
    'PLUGIN_RCS:UNLOAD_MESSAGE_MANUAL': {
      title: 'Unloading',
      message: 'Please remove the current bit, then click <em>"Continue"</em> to proceed or <em>"Abort"</em> to cancel.',
      continueLabel: 'Continue'
    },
    'PLUGIN_RCS:SWAP_MESSAGE_MANUAL': {
      title: 'Tool Change',
      message: 'Please remove the current bit and install {toolNumber} securely, then click <em>"Continue"</em> to proceed or <em>"Abort"</em> to cancel.',
      continueLabel: 'Continue'
    },
    'PLUGIN_RCS:LOAD_MESSAGE': {
      title: 'Loading',
      message: 'Confirm {toolNumber} is placed securely in the pocket and keep hands clear. The spindle will descend to pick up the tool during the load process. Click <em>"Continue"</em> to proceed or <em>"Abort"</em> to cancel.',
      continueLabel: 'Continue'
    },
    'PLUGIN_RCS:UNLOAD_MESSAGE': {
      title: 'Unloading',
      message: 'Ensure the pocket is empty and keep hands clear. The spindle will descend into the pocket during the unload process. Click <em>"Continue"</em> to proceed or <em>"Abort"</em> to cancel.',
      continueLabel: 'Continue'
    }
  };

  ctx.registerEventHandler('ws:cnc-data', async (data) => {
    if (typeof data === 'string') {
      const upperData = data.toUpperCase();
      if (upperData.includes('[MSG') && upperData.includes('PLUGIN_RCS:')) {
        const settings = buildInitialConfig(ctx.getSettings() || {});
        for (const [code, config] of Object.entries(MESSAGE_MAP)) {
          const codePattern = code + '_';
          const codeIndex = upperData.indexOf(codePattern);

          if (codeIndex !== -1) {
            const afterCode = upperData.substring(codeIndex + codePattern.length);
            const toolNumberMatch = afterCode.match(/^(\d+)/);
            const toolNumber = toolNumberMatch ? toolNumberMatch[1] : '';

            const styledToolNumber = `<strong style="color: var(--color-accent);">T${toolNumber}</strong>`;
            const message = config.message.replace('{toolNumber}', styledToolNumber);
            showSafetyWarningDialog(ctx, config.title, message, config.continueLabel, settings.abortEventGcode);
            break;
          } else if (upperData.includes(code)) {
            showSafetyWarningDialog(ctx, config.title, config.message, config.continueLabel, settings.abortEventGcode);
            break;
          }
        }
      }
    }
  });

  // Register onBeforeCommand event handler for M6 interception
  ctx.registerEventHandler('onBeforeCommand', async (commands, context) => {
    const rawSettings = ctx.getSettings() || {};

    // Skip command handling if plugin is not configured
    if (!rawSettings.pocket1) {
      return commands;
    }

    const settings = buildInitialConfig(rawSettings);

    // Handle $H (home) command with conditional TLS
    await handleHomeCommand(commands, context, settings, ctx);

    // Handle $TLS command
    await handleTLSCommand(commands, context, settings, ctx);

    // Handle $POCKET1 command
    handlePocket1Command(commands, settings, ctx);

    // Handle M6 tool change command
    await handleM6Command(commands, context, settings, ctx);

    return commands;
  });

  ctx.registerEventHandler('message', async (data) => {
    if (!data) {
      return;
    }

    if (data.action === 'save') {
      const payload = data.payload || {};
      const sanitized = buildInitialConfig(payload);
      const existing = ctx.getSettings() || {};
      const appSettings = ctx.getAppSettings() || {};
      const resolvedPort = resolveServerPort(existing, appSettings);

      ctx.setSettings({
        ...existing,
        ...sanitized,
        port: resolvedPort
      });

      ctx.log('Manual ToolChanger settings saved');
    }
  });

  ctx.registerToolMenu('Manual Tool Changer', async () => {
    ctx.log('Manual Tool Changer opened');

    const storedSettings = ctx.getSettings() || {};
    const appSettings = ctx.getAppSettings() || {};
    const serverPort = resolveServerPort(storedSettings, appSettings);
    const initialConfig = buildInitialConfig(storedSettings);
    const initialConfigJson = JSON.stringify(initialConfig)
      .replace(/</g, '\\u003c')
      .replace(/>/g, '\\u003e');

    ctx.showDialog(
      'Manual ToolChanger',
      /* html */ `
      <style>
        .rcs-dialog-wrapper {
          display: grid;
          grid-template-rows: auto 1fr auto;
          overflow: hidden;
          width: 900px;
        }

        .rcs-header {
          padding: 10px 30px;
        }

        .rcs-tabs {
          display: flex;
          border-bottom: 1px solid var(--color-border);
          background: var(--color-surface-muted);
          padding: var(--gap-xs) var(--gap-md) 0 var(--gap-md);
          gap: 2px;
          border-left: 1px solid var(--color-border);
          border-right: 1px solid var(--color-border);
          border-top: 1px solid var(--color-border);
        }

        .rcs-tab {
          all: unset;
          display: flex;
          align-items: center;
          gap: var(--gap-xs);
          padding: var(--gap-sm) var(--gap-md);
          background: transparent !important;
          border: none !important;
          border-radius: var(--radius-small) var(--radius-small) 0 0 !important;
          color: var(--color-text-secondary) !important;
          cursor: pointer;
          transition: all 0.2s ease;
          font-size: 0.95rem;
          font-weight: 500;
          margin-top: var(--gap-xs);
          position: relative;
          box-sizing: border-box;
        }

        .rcs-tab:hover {
          background: var(--color-surface) !important;
          color: var(--color-text-primary);
          transform: translateY(-1px);
          filter: none !important;
        }

        .rcs-tab.active {
          background: var(--color-surface) !important;
          color: var(--color-text-primary) !important;
          box-shadow: var(--shadow-elevated);
          border-bottom: 2px solid var(--color-accent) !important;
          filter: none !important;
        }

        .rcs-tab.active::after {
          content: '';
          position: absolute;
          bottom: -1px;
          left: 0;
          right: 0;
          height: 2px;
          background: var(--gradient-accent);
          border-radius: 2px 2px 0 0;
        }

        .rcs-tab:focus-visible {
          outline: 2px solid var(--color-accent);
          outline-offset: 2px;
        }

        .rcs-tab-label {
          font-weight: 600;
        }

        .rcs-content {
          height: 690px;
          display: flex;
          flex-direction: column;
        }

        .rcs-tab-content {
          flex: 1;
          overflow-y: auto;
          padding: 30px;
          padding-top: 20px;
        }

        .rcs-tab-panel {
          display: none;
          height: 100%;
        }

        .rcs-tab-panel.active {
          display: flex;
          flex-direction: column;
        }

        .rcs-container {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 16px;
        }

        .rcs-left-column {
          display: flex;
          flex-direction: column;
          gap: 16px;
        }

        .rcs-right-column {
          display: flex;
          flex-direction: column;
          gap: 16px;
          height: 100%;
        }

        .rcs-axis-card {
          background: color-mix(in srgb, var(--color-surface) 40%, var(--color-surface-muted) 60%);
          border: 1px solid var(--color-border);
          border-radius: var(--radius-small);
          padding: 12px 16px;
          display: flex;
          flex-direction: column;
        }


        .rcs-form-row-single {
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .rcs-form-row-single .rcs-form-label {
          flex: 0 0 auto;
          margin-bottom: 0;
          white-space: nowrap;
        }

        .rcs-form-row-single .rcs-select {
          flex: 0 0 auto;
          width: 100px;
          margin-left: auto;
        }

        .rcs-select {
          padding: 8px 12px;
          padding-right: 24px;
          background: var(--color-surface);
          border: 1px solid var(--color-border);
          border-radius: var(--radius-small);
          color: var(--color-text-primary);
          font-size: 0.9rem;
          font-family: inherit;
          cursor: pointer;
          transition: border-color 0.2s ease;
          appearance: none;
          -webkit-appearance: none;
          -moz-appearance: none;
          background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2393a2b4' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e");
          background-repeat: no-repeat;
          background-position: right 8px center;
          background-size: 16px;
        }

        .rcs-select option {
          text-align: left;
        }

        .rcs-select:hover {
          border-color: var(--color-accent);
        }

        .rcs-select:focus {
          outline: none;
          border-color: var(--color-accent);
          box-shadow: 0 0 0 2px color-mix(in srgb, var(--color-accent) 20%, transparent);
        }

        .rcs-axis-title {
          font-size: 0.85rem;
          font-weight: 600;
          color: var(--color-text-secondary);
          margin-bottom: 8px;
          text-align: center;
        }

        .rcs-axis-values {
          display: flex;
          justify-content: space-around;
          gap: 16px;
          align-items: center;
        }

        .rcs-axis-item {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 4px;
        }

        .rcs-axis-label {
          font-size: 0.75rem;
          font-weight: 700;
          color: var(--color-text-secondary);
          text-transform: uppercase;
        }

        .rcs-axis-value {
          font-family: 'JetBrains Mono', monospace;
          font-size: 1rem;
          font-weight: 600;
          color: var(--color-accent);
        }

        .rcs-form-row {
          display: flex;
          justify-content: center;
          gap: 24px;
        }

        .rcs-form-row-2col {
          display: flex;
          justify-content: center;
          gap: 24px;
        }

        .rcs-form-row-3col {
          display: flex;
          justify-content: center;
          gap: 24px;
        }

        .rcs-form-group {
          display: flex;
          flex-direction: column;
          gap: 8px;
          align-items: center;
        }

        .rcs-form-group-inline {
          display: flex;
          flex-direction: row;
          align-items: center;
          gap: 6px;
        }

        .rcs-form-group-inline .rcs-form-label {
          font-size: 0.85rem;
          min-width: auto;
        }

        .rcs-form-group-inline .rcs-input {
          width: 90px;
        }

        .rcs-form-group-horizontal {
          display: flex;
          flex-direction: row;
          gap: 8px;
          align-items: center;
        }

        .rcs-form-group-horizontal .rcs-form-label {
          white-space: nowrap;
          flex-shrink: 1;
          min-width: 0;
        }

        .rcs-form-group-horizontal .rcs-input {
          width: 100px;
          flex-shrink: 0;
          min-width: 100px;
          margin-left: auto;
        }

        .rcs-form-label {
          font-size: 0.9rem;
          font-weight: 600;
          color: var(--color-text-primary);
        }

        .rcs-input {
          padding: 8px 12px;
          border: 1px solid var(--color-border);
          border-radius: var(--radius-small);
          background: var(--color-surface);
          color: var(--color-text-primary);
          font-size: 0.9rem;
          text-align: center;
          width: 100%;
          max-width: 120px;
        }

        .rcs-input:focus {
          outline: none;
          border-color: var(--color-accent);
        }

        .rcs-input:disabled {
          opacity: 0.4;
          cursor: not-allowed;
          background: var(--color-surface-muted);
        }

        .rcs-control-group {
          background: color-mix(in srgb, var(--color-surface) 40%, var(--color-surface-muted) 60%);
          border: 1px solid var(--color-border);
          border-radius: var(--radius-small);
          padding: 16px;
          display: flex;
          flex-direction: column;
          gap: 16px;
          flex: 1;
          height: 100%;
        }

        .rcs-control-group .rcs-axis-values {
          display: flex;
          justify-content: space-around;
          gap: 16px;
          align-items: center;
          margin-bottom: 8px;
        }

        .rcs-control-group nc-step-control {
          width: 100%;
          display: flex;
          justify-content: center;
          transform: scale(0.95);
        }

        .rcs-control-group nc-jog-control {
          width: 100%;
          display: flex;
          justify-content: center;
        }

        .rcs-pocket-group {
          background: color-mix(in srgb, var(--color-surface) 40%, var(--color-surface-muted) 60%);
          border: 1px solid var(--color-border);
          border-radius: var(--radius-small);
          padding: 16px;
          display: flex;
          flex-direction: column;
          gap: 8px;
        }

        .rcs-pocket-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 12px;
        }

        .rcs-pocket-header-left {
          display: flex;
          align-items: center;
          gap: 12px;
          flex: 1;
          justify-content: center;
        }

        .rcs-pocket-title {
          font-size: 0.95rem;
          font-weight: 600;
          color: var(--color-text-primary);
        }

        .rcs-toggle-switch {
          position: relative;
          width: 56px;
          height: 28px;
          background: var(--color-border);
          border-radius: 14px;
          cursor: pointer;
          transition: background 0.3s ease;
        }

        .rcs-toggle-switch.active {
          background: var(--gradient-accent);
        }

        .rcs-toggle-switch-knob {
          position: absolute;
          top: 2px;
          left: 2px;
          width: 24px;
          height: 24px;
          background: white;
          border-radius: 50%;
          transition: transform 0.3s ease;
        }

        .rcs-toggle-switch.active .rcs-toggle-switch-knob {
          transform: translateX(28px);
        }

        .rcs-toggle-row {
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .rcs-toggle-row .rcs-toggle-switch {
          margin-left: auto;
        }

        .rcs-toggle-label {
          font-size: 0.9rem;
          font-weight: 500;
          color: var(--color-text-primary);
        }

        .rcs-toggle-row.disabled {
          opacity: 0.4;
          pointer-events: none;
        }

        .rcs-button {
          padding: 4px 16px;
          border: none;
          border-radius: var(--radius-small);
          font-weight: 500;
          font-size: 0.8rem;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .rcs-button-grab {
          background: var(--gradient-accent);
          color: white;
          height: 32px;
          display: flex;
          align-items: center;
          justify-content: center;
        }

        .rcs-button-grab:hover {
          opacity: 0.9;
        }

        .rcs-button-grab:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .rcs-footer {
          display: flex;
          justify-content: center;
          gap: 12px;
          padding: 20px 30px;
          border-top: 1px solid var(--color-border);
        }

        .rcs-button-secondary {
          background: var(--color-surface-muted);
          color: var(--color-text-primary);
        }

        .rcs-button-secondary:hover {
          background: var(--color-surface);
        }

        .rcs-button-primary {
          background: var(--gradient-accent);
          color: white;
        }

        .rcs-button-primary:hover {
          opacity: 0.9;
        }

        @keyframes rcs-glow {
          0%, 100% { box-shadow: 0 0 5px var(--color-accent); }
          50% { box-shadow: 0 0 20px var(--color-accent); }
        }

        .rcs-button-saved {
          animation: rcs-glow 0.5s ease-in-out;
        }

        .rcs-button-busy {
          opacity: 0.6;
          cursor: wait;
        }

        .rcs-events-container {
          display: flex;
          flex-direction: column;
          gap: 16px;
          height: 100%;
        }

        .rcs-event-group {
          background: color-mix(in srgb, var(--color-surface) 40%, var(--color-surface-muted) 60%);
          border: 1px solid var(--color-border);
          border-radius: var(--radius-small);
          padding: 16px;
          display: flex;
          flex-direction: column;
          gap: 8px;
          flex: 1;
          min-height: 0;
        }

        .rcs-event-description {
          font-size: 0.8rem;
          color: var(--color-text-secondary);
          margin-bottom: 4px;
        }

        .rcs-monaco-container {
          flex: 1;
          min-height: 80px;
          border: 1px solid var(--color-border);
          border-radius: var(--radius-small);
          overflow: hidden;
        }

        /* Tooltip styles */
        .rcs-tooltip-wrapper {
          position: relative;
          display: inline-flex;
          align-items: center;
        }

        .rcs-tooltip-trigger {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          width: 18px;
          height: 18px;
          margin-left: 2px;
          background: var(--color-accent);
          color: white;
          border-radius: 50%;
          font-size: 11px;
          font-weight: 600;
          cursor: help;
          user-select: none;
          flex-shrink: 0;
        }

        .rcs-tooltip-trigger:hover,
        .rcs-tooltip-trigger:focus {
          background: var(--color-accent-hover, var(--color-accent));
          outline: none;
        }

        .rcs-tooltip-content {
          position: absolute;
          top: calc(100% + 12px);
          left: 50%;
          transform: translateX(-50%);
          width: 280px;
          padding: 12px 14px;
          background: var(--color-accent);
          color: white;
          border-radius: var(--radius-small);
          font-size: 0.85rem;
          line-height: 1.5;
          z-index: 1000;
          opacity: 0;
          visibility: hidden;
          transition: opacity 0.2s, visibility 0.2s;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
        }

        .rcs-tooltip-content::before {
          content: '';
          position: absolute;
          bottom: 100%;
          left: 50%;
          transform: translateX(-50%);
          border: 8px solid transparent;
          border-bottom-color: var(--color-accent);
        }

        .rcs-tooltip-wrapper:hover .rcs-tooltip-content,
        .rcs-tooltip-wrapper:focus-within .rcs-tooltip-content,
        .rcs-tooltip-content.rcs-tooltip-visible {
          opacity: 1;
          visibility: visible;
        }

        /* Tooltip position adjustments for edge cases */
        .rcs-tooltip-content.rcs-tooltip-left {
          left: 0;
          transform: translateX(0);
        }

        .rcs-tooltip-content.rcs-tooltip-left::before {
          left: 14px;
          transform: translateX(0);
        }

        /* Top placement - use transform for reliable positioning */
        .rcs-tooltip-above {
          position: absolute;
          top: 0;
          left: 50%;
          transform: translateX(-50%) translateY(calc(-100% - 12px));
          width: 280px;
          padding: 12px 14px;
          background: var(--color-accent);
          color: white;
          border-radius: var(--radius-small);
          font-size: 0.85rem;
          line-height: 1.5;
          z-index: 1000;
          opacity: 0;
          visibility: hidden;
          transition: opacity 0.2s, visibility 0.2s;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
        }

        .rcs-tooltip-above::before {
          content: '';
          position: absolute;
          top: 100%;
          left: 50%;
          transform: translateX(-50%);
          border: 8px solid transparent;
          border-top-color: var(--color-accent);
        }

        .rcs-tooltip-wrapper:hover .rcs-tooltip-above,
        .rcs-tooltip-wrapper:focus-within .rcs-tooltip-above,
        .rcs-tooltip-above.rcs-tooltip-visible {
          opacity: 1;
          visibility: visible;
        }

        .rcs-tooltip-above.rcs-tooltip-right {
          left: auto;
          right: 0;
          transform: translateX(0) translateY(calc(-100% - 12px));
        }

        .rcs-tooltip-above.rcs-tooltip-right::before {
          left: auto;
          right: 14px;
          transform: translateX(0);
        }
      </style>

      <div class="rcs-dialog-wrapper">
        <div class="rcs-tabs">
          <button class="rcs-tab active" data-tab="basic">
            <span class="rcs-tab-label">Basic</span>
          </button>
          <button class="rcs-tab" data-tab="events">
            <span class="rcs-tab-label">Events</span>
          </button>
        </div>
        <div class="rcs-content">
          <div class="rcs-tab-content">
            <!-- Basic Tab -->
            <div class="rcs-tab-panel active" id="rcs-tab-basic">
              <div class="rcs-container">
            <!-- Left Column -->
            <div class="rcs-left-column">
              <!-- Tool Setter Location -->
            <div class="rcs-pocket-group">
              <div class="rcs-pocket-header">
                <div class="rcs-pocket-header-left">
                  <span class="rcs-pocket-title">Tool Setter</span>
                  <button type="button" class="rcs-button rcs-button-grab" id="rcs-toolsetter-grab">Grab</button>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-content rcs-tooltip-left">Machine coordinates of your Tool Length Setter (TLS) probe. Use your longest bit, position the spindle directly above the probe center with the tip 1 to 1.5 inches above the setter, then click Grab to capture coordinates.</div>
                  </div>
                </div>
              </div>

              <div class="rcs-form-row-3col">
                <div class="rcs-form-group-inline">
                  <label class="rcs-form-label">X</label>
                  <input type="number" class="rcs-input" id="rcs-toolsetter-x" value="0" step="0.001">
                </div>
                <div class="rcs-form-group-inline">
                  <label class="rcs-form-label">Y</label>
                  <input type="number" class="rcs-input" id="rcs-toolsetter-y" value="0" step="0.001">
                </div>
                <div class="rcs-form-group-inline">
                  <label class="rcs-form-label">Z</label>
                  <input type="number" class="rcs-input" id="rcs-toolsetter-z" value="0" step="0.001">
                </div>
              </div>

              <div class="rcs-form-group-horizontal">
                <label class="rcs-form-label">Seek Distance (mm)</label>
                <div class="rcs-tooltip-wrapper">
                  <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                  <div class="rcs-tooltip-content rcs-tooltip-left">Maximum distance the tool will travel downward while probing for the tool setter. If no contact is made within this distance, the probe fails.</div>
                </div>
                <input type="number" class="rcs-input" id="rcs-seek-distance" value="50" step="1" min="1">
              </div>

              <div class="rcs-form-group-horizontal">
                <label class="rcs-form-label">Seek Feedrate (mm/min)</label>
                <div class="rcs-tooltip-wrapper">
                  <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                  <div class="rcs-tooltip-content rcs-tooltip-left">Speed at which the tool moves downward during probing. Lower values provide more accurate measurements but take longer.</div>
                </div>
                <input type="number" class="rcs-input" id="rcs-seek-feedrate" value="100" step="10" min="1">
              </div>
            </div>

              <!-- Manual ToolChange Location -->
              <div class="rcs-pocket-group">
                <div class="rcs-pocket-header">
                  <div class="rcs-pocket-header-left">
                    <span class="rcs-pocket-title">Manual ToolChange</span>
                    <button type="button" class="rcs-button rcs-button-grab" id="rcs-parking-grab">Grab</button>
                    <div class="rcs-tooltip-wrapper">
                      <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                      <div class="rcs-tooltip-content rcs-tooltip-left">Location where the spindle moves for manual tool changes. Position this where you have easy access to change tools by hand.</div>
                    </div>
                  </div>
                </div>

                <div class="rcs-form-row-2col">
                  <div class="rcs-form-group-inline">
                    <label class="rcs-form-label">X</label>
                    <input type="number" class="rcs-input" id="rcs-parking-x" value="0" step="0.001">
                  </div>
                  <div class="rcs-form-group-inline">
                    <label class="rcs-form-label">Y</label>
                    <input type="number" class="rcs-input" id="rcs-parking-y" value="0" step="0.001">
                  </div>
                </div>
              </div>

              <!-- RapidChangeSolo -->
              <div class="rcs-pocket-group">
                <div class="rcs-pocket-header">
                  <div class="rcs-pocket-header-left">
                    <div class="rcs-toggle-switch" id="rcs-autoswap-toggle">
                      <div class="rcs-toggle-switch-knob"></div>
                    </div>
                    <span class="rcs-pocket-title">RapidChangeSolo</span>
                    <button type="button" class="rcs-button rcs-button-grab" id="rcs-pocket1-grab">Grab</button>
                    <div class="rcs-tooltip-wrapper">
                      <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                      <div class="rcs-tooltip-content rcs-tooltip-left">Enable if using RapidChangeSolo tool holder. Setup: Position spindle over the Solo pocket without rubbing, lower Z until the nut engages the ball bearing inside the pocket, then click Grab. The Grab button automatically adds -5mm to the Z value for proper engagement.</div>
                    </div>
                  </div>
                </div>

                <div class="rcs-form-row">
                  <div class="rcs-form-group-inline">
                    <label class="rcs-form-label">X</label>
                    <input type="number" class="rcs-input" id="rcs-pocket1-x" value="0" step="0.001">
                  </div>
                  <div class="rcs-form-group-inline">
                    <label class="rcs-form-label">Y</label>
                    <input type="number" class="rcs-input" id="rcs-pocket1-y" value="0" step="0.001">
                  </div>
                  <div class="rcs-form-group-inline">
                    <label class="rcs-form-label">Z</label>
                    <input type="number" class="rcs-input" id="rcs-zengagement" value="-50" step="0.001">
                  </div>
                </div>

                <div class="rcs-toggle-row" id="rcs-pause-before-unload-row">
                  <span class="rcs-toggle-label">Pause before Unload</span>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-content">When enabled, the system will pause and wait for user confirmation before unloading the current tool. Useful for verifying the pocket is empty before dropping the tool.</div>
                  </div>
                  <div class="rcs-toggle-switch active" id="rcs-pause-before-unload-toggle">
                    <div class="rcs-toggle-switch-knob"></div>
                  </div>
                </div>

                <div class="rcs-toggle-row" id="rcs-wait-for-spindle-row">
                  <span class="rcs-toggle-label">Wait for Spindle</span>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-content">When enabled, the system waits for the spindle to reach the target RPM before proceeding with the tool change sequence. Recommended for reliable operation.</div>
                  </div>
                  <div class="rcs-toggle-switch active" id="rcs-wait-for-spindle-toggle">
                    <div class="rcs-toggle-switch-knob"></div>
                  </div>
                </div>

                <div class="rcs-form-group-horizontal" id="rcs-unload-rpm-row">
                  <label class="rcs-form-label">Unload RPM</label>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-content">Spindle speed (CCW rotation) used when unloading. The reverse rotation unthreads the nut/collet/bit to drop it into the holder pocket.</div>
                  </div>
                  <input type="number" class="rcs-input" id="rcs-unload-rpm" value="1500" min="500" max="3000" step="1">
                </div>

                <div class="rcs-form-group-horizontal" id="rcs-load-rpm-row">
                  <label class="rcs-form-label">Load RPM</label>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-above">Spindle speed (CW rotation) used when loading. The forward rotation threads the nut/collet/bit to pick it up from the holder pocket.</div>
                  </div>
                  <input type="number" class="rcs-input" id="rcs-load-rpm" value="1200" min="500" max="3000" step="1">
                </div>

                <div class="rcs-form-group-horizontal" id="rcs-z-retreat-row">
                  <label class="rcs-form-label">Retract (mm)</label>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-above">Distance for tightening/untightening retries. The spindle plunges 3 times while spinning to ensure proper threading or to retry if it doesn't release on the first attempt.</div>
                  </div>
                  <input type="number" class="rcs-input" id="rcs-z-retreat" value="12" min="5" max="20" step="1">
                </div>

              </div>
            </div>

            <!-- Right Column -->
            <div class="rcs-right-column">
              <!-- Controls Group with Machine Coordinates -->
              <div class="rcs-control-group">
                <div class="rcs-pocket-header">
                  <div class="rcs-pocket-header-left">
                    <span class="rcs-pocket-title">Machine Coordinates</span>
                  </div>
                </div>
                <div class="rcs-axis-values">
                  <div class="rcs-axis-item">
                    <span class="rcs-axis-label">X</span>
                    <span class="rcs-axis-value" id="rcs-axis-x">0.000</span>
                  </div>
                  <div class="rcs-axis-item">
                    <span class="rcs-axis-label">Y</span>
                    <span class="rcs-axis-value" id="rcs-axis-y">0.000</span>
                  </div>
                  <div class="rcs-axis-item">
                    <span class="rcs-axis-label">Z</span>
                    <span class="rcs-axis-value" id="rcs-axis-z">0.000</span>
                  </div>
                </div>
                <nc-step-control></nc-step-control>
                <nc-jog-control></nc-jog-control>
              </div>

              <!-- Tool Settings -->
              <div class="rcs-pocket-group">
                <div class="rcs-pocket-header">
                  <div class="rcs-pocket-header-left">
                    <span class="rcs-pocket-title">Tool Settings</span>
                  </div>
                </div>

                <div class="rcs-form-row-single">
                  <label class="rcs-form-label">Number of Tools</label>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-above rcs-tooltip-right">Virtual magazine size. Determines how many tool buttons are displayed on the main interface.</div>
                  </div>
                  <select class="rcs-select" id="rcs-number-of-tools">
                  </select>
                </div>

                <div class="rcs-toggle-row">
                  <span class="rcs-toggle-label">Show Command</span>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-above rcs-tooltip-right">When enabled, displays the generated G-code macro command (e.g., o100 call) that you can copy into your post processor or G-code files.</div>
                  </div>
                  <div class="rcs-toggle-switch" id="rcs-show-macro-command-toggle">
                    <div class="rcs-toggle-switch-knob"></div>
                  </div>
                </div>

                <div class="rcs-toggle-row">
                  <span class="rcs-toggle-label">Perform TLS after first HOME</span>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-above rcs-tooltip-right">When enabled, automatically performs a Tool Length Setter probe cycle after the first homing operation. Ensures accurate tool offset before starting a job.</div>
                  </div>
                  <div class="rcs-toggle-switch" id="rcs-perform-tls-after-home-toggle">
                    <div class="rcs-toggle-switch-knob"></div>
                  </div>
                </div>

                <div class="rcs-form-row-single">
                  <label class="rcs-form-label">Switch Aux during TLS</label>
                  <div class="rcs-tooltip-wrapper">
                    <span class="rcs-tooltip-trigger" tabindex="0">?</span>
                    <div class="rcs-tooltip-above rcs-tooltip-right">Select an auxiliary output to activate during tool length probing. Useful for controlling dust collection, air blast, or probe activation relays.</div>
                  </div>
                  <select class="rcs-select" id="rcs-tls-aux-output">
                    <option value="-1">Disabled</option>
                  </select>
                </div>
              </div>
            </div>
          </div>
            </div>

            <!-- Events Tab -->
            <div class="rcs-tab-panel" id="rcs-tab-events">
              <div class="rcs-events-container">
                <div class="rcs-event-group">
                  <label class="rcs-form-label" style="text-align: left; font-size: 14px;">Pre Tool Change Event</label>
                  <div class="rcs-event-description">G-code to execute at the very start of the tool change sequence, before any other operations.</div>
                  <div class="rcs-monaco-container" id="rcs-pre-tool-change-editor"></div>
                </div>

                <div class="rcs-event-group">
                  <label class="rcs-form-label" style="text-align: left; font-size: 14px;">Post Tool Change Event</label>
                  <div class="rcs-event-description">G-code to execute at the very end of the tool change sequence, after all operations complete.</div>
                  <div class="rcs-monaco-container" id="rcs-post-tool-change-editor"></div>
                </div>

                <div class="rcs-event-group">
                  <label class="rcs-form-label" style="text-align: left; font-size: 14px;">Abort Event</label>
                  <div class="rcs-event-description">G-code to execute when the abort button is pressed during tool change. Runs before soft reset.</div>
                  <div class="rcs-monaco-container" id="rcs-abort-event-editor"></div>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div class="rcs-footer">
          <button type="button" class="rcs-button rcs-button-secondary" id="rcs-close-btn">Close</button>
          <button type="button" class="rcs-button rcs-button-primary" id="rcs-save-btn">Save</button>
        </div>
      </div>

      <script>
        (function() {
          const BASE_URL = '';
          const POCKET_PREFIX = 'pocket1';
          const TOOLSETTER_PREFIX = 'toolsetter';
          const PARKING_PREFIX = 'parking';

          const initialConfig = JSON.parse('${initialConfigJson}');

          // Monaco editor instances
          let preToolChangeEditor = null;
          let postToolChangeEditor = null;
          let abortEventEditor = null;

          // Tab switching
          const tabs = document.querySelectorAll('.rcs-tab');
          const tabPanels = document.querySelectorAll('.rcs-tab-panel');

          tabs.forEach(tab => {
            tab.addEventListener('click', () => {
              const targetTab = tab.dataset.tab;

              tabs.forEach(t => t.classList.remove('active'));
              tabPanels.forEach(p => p.classList.remove('active'));

              tab.classList.add('active');
              document.getElementById('rcs-tab-' + targetTab).classList.add('active');

              // Layout Monaco editors when Events tab is shown
              if (targetTab === 'events') {
                if (preToolChangeEditor) preToolChangeEditor.layout();
                if (postToolChangeEditor) postToolChangeEditor.layout();
                if (abortEventEditor) abortEventEditor.layout();
              }
            });
          });

          // Tooltip touch support - tap to toggle
          const tooltipWrappers = document.querySelectorAll('.rcs-tooltip-wrapper');
          let activeTooltip = null;

          const hideAllTooltips = () => {
            document.querySelectorAll('.rcs-tooltip-content, .rcs-tooltip-above').forEach(t => {
              t.classList.remove('rcs-tooltip-visible');
            });
            activeTooltip = null;
          };

          tooltipWrappers.forEach(wrapper => {
            const trigger = wrapper.querySelector('.rcs-tooltip-trigger');
            const content = wrapper.querySelector('.rcs-tooltip-content') || wrapper.querySelector('.rcs-tooltip-above');
            if (!trigger || !content) return;

            trigger.addEventListener('click', (e) => {
              e.preventDefault();
              e.stopPropagation();
              if (activeTooltip === content) {
                content.classList.remove('rcs-tooltip-visible');
                activeTooltip = null;
              } else {
                hideAllTooltips();
                content.classList.add('rcs-tooltip-visible');
                activeTooltip = content;
              }
            });
          });

          document.addEventListener('click', (e) => {
            if (activeTooltip && !e.target.closest('.rcs-tooltip-wrapper')) {
              hideAllTooltips();
            }
          });

          // Initialize Monaco editors
          const isLightTheme = () => document.body.classList.contains('theme-light');
          const getMonacoTheme = () => isLightTheme() ? 'gcode-light' : 'gcode-dark';
          const monacoOptions = {
            language: 'gcode',
            theme: getMonacoTheme(),
            minimap: { enabled: false },
            lineNumbers: 'on',
            scrollBeyondLastLine: false,
            automaticLayout: false,
            fontSize: 12,
            tabSize: 2
          };

          if (typeof monaco !== 'undefined') {
            const preToolChangeContainer = document.getElementById('rcs-pre-tool-change-editor');
            if (preToolChangeContainer) {
              preToolChangeEditor = monaco.editor.create(preToolChangeContainer, {
                ...monacoOptions,
                value: initialConfig.preToolChangeGcode || ''
              });
            }

            const postToolChangeContainer = document.getElementById('rcs-post-tool-change-editor');
            if (postToolChangeContainer) {
              postToolChangeEditor = monaco.editor.create(postToolChangeContainer, {
                ...monacoOptions,
                value: initialConfig.postToolChangeGcode || ''
              });
            }

            const abortEventContainer = document.getElementById('rcs-abort-event-editor');
            if (abortEventContainer) {
              abortEventEditor = monaco.editor.create(abortEventContainer, {
                ...monacoOptions,
                value: initialConfig.abortEventGcode || ''
              });
            }

            // Watch for theme changes
            const themeObserver = new MutationObserver(() => {
              monaco.editor.setTheme(getMonacoTheme());
            });
            themeObserver.observe(document.body, { attributes: true, attributeFilter: ['class'] });
          }

          const getInput = (id) => document.getElementById(id);
          const formatCoordinate = (value) => (Number.isFinite(value) ? value.toFixed(3) : '0.000');

          const parseCoordinateString = (raw) => {
            if (typeof raw === 'string' && raw.length > 0) {
              const parts = raw.split(',').map((part) => Number.parseFloat(part.trim()));
              if (parts.length >= 2 && parts.every(Number.isFinite)) {
                return { x: parts[0], y: parts[1], z: parts[2] };
              }
            }
            if (Array.isArray(raw) && raw.length >= 2) {
              const [x, y, z] = raw;
              if ([x, y].every(Number.isFinite)) {
                return { x, y, z };
              }
            }
            if (raw && typeof raw === 'object') {
              const { x, y, z } = raw;
              if ([x, y].every(Number.isFinite)) {
                return { x, y, z };
              }
            }
            return null;
          };

          const extractCoordinatesFromPayload = (payload) => {
            if (!payload || typeof payload !== 'object') {
              return null;
            }

            const nestedKeys = ['machineState', 'lastStatus', 'statusReport'];
            for (let i = 0; i < nestedKeys.length; i += 1) {
              const key = nestedKeys[i];
              if (payload[key] && typeof payload[key] === 'object') {
                const nestedCoords = extractCoordinatesFromPayload(payload[key]);
                if (nestedCoords) {
                  return nestedCoords;
                }
              }
            }

            const candidates = [
              payload.machineCoords,
              payload.MPos,
              payload.MPOS,
              payload.mpos,
              payload.machinePosition
            ];

            for (let i = 0; i < candidates.length; i += 1) {
              const coords = parseCoordinateString(candidates[i]);
              if (coords) {
                return coords;
              }
            }

            return null;
          };

          const applyInitialSettings = () => {
            if (initialConfig.pocket1) {
              getInput('rcs-pocket1-x').value = initialConfig.pocket1.x || 0;
              getInput('rcs-pocket1-y').value = initialConfig.pocket1.y || 0;
            }

            if (initialConfig.toolSetter) {
              getInput('rcs-toolsetter-x').value = initialConfig.toolSetter.x || 0;
              getInput('rcs-toolsetter-y').value = initialConfig.toolSetter.y || 0;
              getInput('rcs-toolsetter-z').value = initialConfig.toolSetter.z || 0;
            }

            if (initialConfig.parking) {
              getInput('rcs-parking-x').value = initialConfig.parking.x || 0;
              getInput('rcs-parking-y').value = initialConfig.parking.y || 0;
            }

            getInput('rcs-zengagement').value = initialConfig.zEngagement || -50;
            getInput('rcs-seek-distance').value = initialConfig.seekDistance || 50;
            getInput('rcs-seek-feedrate').value = initialConfig.seekFeedrate || 100;
            getInput('rcs-number-of-tools').value = initialConfig.numberOfTools || 1;
            getInput('rcs-unload-rpm').value = initialConfig.unloadRpm || 1500;
            getInput('rcs-load-rpm').value = initialConfig.loadRpm || 1200;
            getInput('rcs-z-retreat').value = initialConfig.zRetreat || 12;

            const autoSwapToggle = document.getElementById('rcs-autoswap-toggle');
            if (autoSwapToggle) {
              if (initialConfig.autoSwap) {
                autoSwapToggle.classList.add('active');
              } else {
                autoSwapToggle.classList.remove('active');
              }
            }

            const pauseBeforeUnloadToggle = document.getElementById('rcs-pause-before-unload-toggle');
            if (pauseBeforeUnloadToggle) {
              if (initialConfig.pauseBeforeUnload) {
                pauseBeforeUnloadToggle.classList.add('active');
              } else {
                pauseBeforeUnloadToggle.classList.remove('active');
              }
            }

            const waitForSpindleToggle = document.getElementById('rcs-wait-for-spindle-toggle');
            if (waitForSpindleToggle) {
              if (initialConfig.waitForSpindle) {
                waitForSpindleToggle.classList.add('active');
              } else {
                waitForSpindleToggle.classList.remove('active');
              }
            }

            const showMacroCommandToggle = document.getElementById('rcs-show-macro-command-toggle');
            if (showMacroCommandToggle) {
              if (initialConfig.showMacroCommand) {
                showMacroCommandToggle.classList.add('active');
              } else {
                showMacroCommandToggle.classList.remove('active');
              }
            }

            const performTlsAfterHomeToggle = document.getElementById('rcs-perform-tls-after-home-toggle');
            if (performTlsAfterHomeToggle) {
              if (initialConfig.performTlsAfterHome) {
                performTlsAfterHomeToggle.classList.add('active');
              } else {
                performTlsAfterHomeToggle.classList.remove('active');
              }
            }

            // tlsAuxOutput is set after dropdown is populated
          };

          const grabCoordinates = async (prefix) => {
            try {
              const response = await fetch(BASE_URL + '/api/server-state');
              if (!response.ok) {
                throw new Error('Failed to fetch server state: ' + response.status);
              }

              const state = await response.json();
              const coords = extractCoordinatesFromPayload(state);

              if (!coords) {
                throw new Error('No coordinates available in server state');
              }

              getInput(\`rcs-\${prefix}-x\`).value = (coords.x || 0).toFixed(3);
              getInput(\`rcs-\${prefix}-y\`).value = (coords.y || 0).toFixed(3);

              if (prefix === POCKET_PREFIX) {
                const zEngagement = (coords.z || 0) - 5;
                getInput('rcs-zengagement').value = zEngagement.toFixed(3);
              }

              if (prefix === TOOLSETTER_PREFIX) {
                getInput(\`rcs-\${prefix}-z\`).value = (coords.z || 0).toFixed(3);
              }

            } catch (error) {
              console.error('[ManualToolChanger] Failed to grab coordinates:', error);
            }
          };

          const registerButton = (prefix, buttonId) => {
            const button = getInput(buttonId);
            if (button) {
              button.addEventListener('click', async function() {
                if (button.disabled) return;
                button.disabled = true;

                try {
                  await grabCoordinates(prefix);
                } finally {
                  setTimeout(() => { button.disabled = false; }, 500);
                }
              });
            }
          };

          const gatherFormData = () => {
            const autoSwapToggle = document.getElementById('rcs-autoswap-toggle');
            const pauseBeforeUnloadToggle = document.getElementById('rcs-pause-before-unload-toggle');
            const waitForSpindleToggle = document.getElementById('rcs-wait-for-spindle-toggle');
            const showMacroCommandToggle = document.getElementById('rcs-show-macro-command-toggle');
            const performTlsAfterHomeToggle = document.getElementById('rcs-perform-tls-after-home-toggle');

            return {
              pocket1: {
                x: parseFloat(getInput('rcs-pocket1-x').value) || 0,
                y: parseFloat(getInput('rcs-pocket1-y').value) || 0
              },
              toolSetter: {
                x: parseFloat(getInput('rcs-toolsetter-x').value) || 0,
                y: parseFloat(getInput('rcs-toolsetter-y').value) || 0,
                z: parseFloat(getInput('rcs-toolsetter-z').value) || 0
              },
              parking: {
                x: parseFloat(getInput('rcs-parking-x').value) || 0,
                y: parseFloat(getInput('rcs-parking-y').value) || 0
              },
              zEngagement: parseFloat(getInput('rcs-zengagement').value) || -50,
              seekDistance: parseFloat(getInput('rcs-seek-distance').value) || 50,
              seekFeedrate: parseFloat(getInput('rcs-seek-feedrate').value) || 100,
              numberOfTools: parseInt(getInput('rcs-number-of-tools').value) || 1,
              unloadRpm: Math.min(3000, Math.max(500, parseInt(getInput('rcs-unload-rpm').value) || 1500)),
              loadRpm: Math.min(3000, Math.max(500, parseInt(getInput('rcs-load-rpm').value) || 1200)),
              zRetreat: Math.min(20, Math.max(5, parseInt(getInput('rcs-z-retreat').value) || 12)),
              autoSwap: autoSwapToggle ? autoSwapToggle.classList.contains('active') : false,
              pauseBeforeUnload: pauseBeforeUnloadToggle ? pauseBeforeUnloadToggle.classList.contains('active') : true,
              waitForSpindle: waitForSpindleToggle ? waitForSpindleToggle.classList.contains('active') : true,
              showMacroCommand: showMacroCommandToggle ? showMacroCommandToggle.classList.contains('active') : false,
              performTlsAfterHome: performTlsAfterHomeToggle ? performTlsAfterHomeToggle.classList.contains('active') : false,
              tlsAuxOutput: (() => {
                const val = getInput('rcs-tls-aux-output').value;
                if (val === 'M7' || val === 'M8') return val;
                return parseInt(val) || -1;
              })(),
              preToolChangeGcode: preToolChangeEditor ? preToolChangeEditor.getValue() : '',
              postToolChangeGcode: postToolChangeEditor ? postToolChangeEditor.getValue() : '',
              abortEventGcode: abortEventEditor ? abortEventEditor.getValue() : ''
            };
          };

          const closeButton = getInput('rcs-close-btn');
          if (closeButton) {
            closeButton.addEventListener('click', function() {
              window.postMessage({ type: 'close-plugin-dialog' }, '*');
            });
          };

          const saveButton = getInput('rcs-save-btn');
          if (saveButton) {
            saveButton.addEventListener('click', async function() {
              if (saveButton.disabled) return;

              saveButton.disabled = true;
              saveButton.classList.add('rcs-button-busy');

              const payload = gatherFormData();

              try {
                const pluginResponse = await fetch(BASE_URL + '/api/plugins/com.ncsender.manualtoolchange/settings', {
                  method: 'PUT',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify(payload)
                });

                if (!pluginResponse.ok) {
                  throw new Error('Failed to save plugin settings: ' + pluginResponse.status);
                }

                // Update tool settings to enable manual and TLS
                const settingsResponse = await fetch(BASE_URL + '/api/settings', {
                  method: 'PATCH',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({
                    tool: {
                      count: payload.numberOfTools || 1,
                      source: 'com.ncsender.manualtoolchange',
                      manual: true,
                      tls: true
                    }
                  })
                });

                if (!settingsResponse.ok) {
                  throw new Error('Failed to update tool settings: ' + settingsResponse.status);
                }

                saveButton.textContent = 'Saved';
                saveButton.classList.add('rcs-button-saved');
                saveButton.classList.remove('rcs-button-busy');
                saveButton.disabled = false;

                setTimeout(function() {
                  saveButton.textContent = 'Save';
                  saveButton.classList.remove('rcs-button-saved');
                }, 2000);
              } catch (error) {
                console.error('[ManualToolChanger] Failed to save settings:', error);
                saveButton.disabled = false;
                saveButton.classList.remove('rcs-button-busy');
              }
            });
          }

          const updateAxisDisplay = (coords) => {
            if (!coords) return;

            const axisX = document.getElementById('rcs-axis-x');
            const axisY = document.getElementById('rcs-axis-y');
            const axisZ = document.getElementById('rcs-axis-z');

            if (axisX && coords.x !== undefined) axisX.textContent = formatCoordinate(coords.x);
            if (axisY && coords.y !== undefined) axisY.textContent = formatCoordinate(coords.y);
            if (axisZ && coords.z !== undefined) axisZ.textContent = formatCoordinate(coords.z);
          };

          const handleServerStateUpdate = (event) => {
            if (!event.data || event.data.type !== 'server-state-update') return;

            const coords = extractCoordinatesFromPayload(event.data.state);
            if (coords) {
              updateAxisDisplay(coords);
            }
          };

          const fetchInitialCoordinates = async () => {
            try {
              const response = await fetch(BASE_URL + '/api/server-state');
              if (!response.ok) return;

              const state = await response.json();
              const coords = extractCoordinatesFromPayload(state);

              if (coords) {
                updateAxisDisplay(coords);
              }
            } catch (error) {
              console.error('[ManualToolChanger] Failed to fetch initial coordinates:', error);
            }
          };

          const populateAuxOutputs = async () => {
            try {
              // Fetch both server state (for outputPins count) and settings (for custom names)
              const [stateResponse, settingsResponse] = await Promise.all([
                fetch(BASE_URL + '/api/server-state'),
                fetch(BASE_URL + '/api/settings')
              ]);

              if (!stateResponse.ok) return;

              const state = await stateResponse.json();
              const settings = settingsResponse.ok ? await settingsResponse.json() : {};
              const outputPins = state?.machineState?.outputPins || state?.outputPins || 0;
              const auxOutputsConfig = settings?.auxOutputs || [];

              // Build a map of G-code command to custom name
              const nameMap = {};
              for (const output of auxOutputsConfig) {
                if (output.on && output.name) {
                  nameMap[output.on.toUpperCase()] = output.name;
                }
              }

              const select = getInput('rcs-tls-aux-output');
              if (!select) return;

              // Add M7 and M8 coolant options (check for custom names)
              const m8Option = document.createElement('option');
              m8Option.value = 'M8';
              m8Option.textContent = nameMap['M8'] || 'M8';
              select.appendChild(m8Option);

              const m7Option = document.createElement('option');
              m7Option.value = 'M7';
              m7Option.textContent = nameMap['M7'] || 'M7';
              select.appendChild(m7Option);

              // Add options for each available aux output
              for (let i = 0; i < outputPins; i++) {
                const option = document.createElement('option');
                const gcode = 'M64 P' + i;
                option.value = i;
                option.textContent = nameMap[gcode.toUpperCase()] || gcode;
                select.appendChild(option);
              }

              // Set initial value
              select.value = initialConfig.tlsAuxOutput ?? -1;
            } catch (error) {
              console.error('[ManualToolChanger] Failed to populate aux outputs:', error);
            }
          };

          const autoSwapToggle = document.getElementById('rcs-autoswap-toggle');
          const pauseBeforeUnloadRow = document.getElementById('rcs-pause-before-unload-row');
          const waitForSpindleRow = document.getElementById('rcs-wait-for-spindle-row');
          const unloadRpmRow = document.getElementById('rcs-unload-rpm-row');
          const loadRpmRow = document.getElementById('rcs-load-rpm-row');
          const zRetreatRow = document.getElementById('rcs-z-retreat-row');

          const updateRapidChangeState = () => {
            const isEnabled = autoSwapToggle && autoSwapToggle.classList.contains('active');

            // Update X, Y, Z inputs and Grab button
            const pocket1X = getInput('rcs-pocket1-x');
            const pocket1Y = getInput('rcs-pocket1-y');
            const zEngagement = getInput('rcs-zengagement');
            const grabBtn = getInput('rcs-pocket1-grab');
            const unloadRpm = getInput('rcs-unload-rpm');
            const loadRpm = getInput('rcs-load-rpm');
            const zRetreat = getInput('rcs-z-retreat');

            [pocket1X, pocket1Y, zEngagement, grabBtn, unloadRpm, loadRpm, zRetreat].forEach(el => {
              if (el) el.disabled = !isEnabled;
            });

            // Update Pause before Unload row
            if (pauseBeforeUnloadRow) {
              if (isEnabled) {
                pauseBeforeUnloadRow.classList.remove('disabled');
              } else {
                pauseBeforeUnloadRow.classList.add('disabled');
              }
            }

            // Update Wait for Spindle row
            if (waitForSpindleRow) {
              if (isEnabled) {
                waitForSpindleRow.classList.remove('disabled');
              } else {
                waitForSpindleRow.classList.add('disabled');
              }
            }

            // Update Unload RPM row
            if (unloadRpmRow) {
              if (isEnabled) {
                unloadRpmRow.classList.remove('disabled');
              } else {
                unloadRpmRow.classList.add('disabled');
              }
            }

            // Update Load RPM row
            if (loadRpmRow) {
              if (isEnabled) {
                loadRpmRow.classList.remove('disabled');
              } else {
                loadRpmRow.classList.add('disabled');
              }
            }

            // Update Z Retreat row
            if (zRetreatRow) {
              if (isEnabled) {
                zRetreatRow.classList.remove('disabled');
              } else {
                zRetreatRow.classList.add('disabled');
              }
            }
          };

          if (autoSwapToggle) {
            autoSwapToggle.addEventListener('click', function() {
              const isCurrentlyActive = autoSwapToggle.classList.contains('active');

              // If turning ON, show warning
              if (!isCurrentlyActive) {
                showAutoSwapWarning(function(confirmed) {
                  if (confirmed) {
                    autoSwapToggle.classList.add('active');
                    updateRapidChangeState();
                  }
                });
              } else {
                // Turning OFF - no warning needed
                autoSwapToggle.classList.remove('active');
                updateRapidChangeState();
              }
            });
          }

          function showAutoSwapWarning(callback) {
            // Create overlay
            const overlay = document.createElement('div');
            overlay.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0, 0, 0, 0.7); display: flex; align-items: center; justify-content: center; z-index: 10000;';

            // Create warning container
            const container = document.createElement('div');
            container.style.cssText = 'background: var(--color-surface); border-radius: var(--radius-medium); padding: 32px; max-width: 500px; box-shadow: 0 20px 60px rgba(0, 0, 0, 0.5);';

            // Create header
            const header = document.createElement('div');
            header.style.cssText = 'font-size: 1.5rem; font-weight: 600; color: var(--color-text-primary); margin-bottom: 24px; text-align: center;';
            header.textContent = 'RapidChangeSolo';

            // Create message
            const message = document.createElement('div');
            message.style.cssText = 'font-size: 1rem; line-height: 1.6; color: var(--color-text-primary); background: color-mix(in srgb, var(--color-warning) 15%, transparent); border: 2px solid var(--color-warning); border-radius: var(--radius-small); padding: 16px; margin-bottom: 16px;';
            message.innerHTML = 'When enabled, the spindle will automatically spin during tool swapping operations. This is required for the RapidChangeSolo mechanism to properly engage and release tools.<br><br><strong>Only enable this if you are using a RapidChangeSolo tool holder system.</strong>';

            // Create buttons container
            const actions = document.createElement('div');
            actions.style.cssText = 'display: flex; justify-content: center; gap: 16px; margin-top: 8px;';

            // Cancel button
            const cancelBtn = document.createElement('button');
            cancelBtn.textContent = 'Cancel';
            cancelBtn.style.cssText = 'padding: 12px 32px; border: none; border-radius: var(--radius-small); font-weight: 600; font-size: 1rem; cursor: pointer; background: var(--color-surface-muted); color: var(--color-text-primary); min-width: 120px;';
            cancelBtn.onclick = function() {
              document.body.removeChild(overlay);
              callback(false);
            };

            // Confirm button
            const confirmBtn = document.createElement('button');
            confirmBtn.textContent = 'I Understand, Enable';
            confirmBtn.style.cssText = 'padding: 12px 32px; border: none; border-radius: var(--radius-small); font-weight: 600; font-size: 1rem; cursor: pointer; background: var(--gradient-accent); color: white; min-width: 120px;';
            confirmBtn.onclick = function() {
              document.body.removeChild(overlay);
              callback(true);
            };

            // Assemble
            actions.appendChild(cancelBtn);
            actions.appendChild(confirmBtn);
            container.appendChild(header);
            container.appendChild(message);
            container.appendChild(actions);
            overlay.appendChild(container);
            document.body.appendChild(overlay);
          }

          const pauseBeforeUnloadToggle = document.getElementById('rcs-pause-before-unload-toggle');
          if (pauseBeforeUnloadToggle) {
            pauseBeforeUnloadToggle.addEventListener('click', function() {
              pauseBeforeUnloadToggle.classList.toggle('active');
            });
          }

          const waitForSpindleToggle = document.getElementById('rcs-wait-for-spindle-toggle');
          if (waitForSpindleToggle) {
            waitForSpindleToggle.addEventListener('click', function() {
              waitForSpindleToggle.classList.toggle('active');
            });
          }

          const showMacroCommandToggle = document.getElementById('rcs-show-macro-command-toggle');
          if (showMacroCommandToggle) {
            showMacroCommandToggle.addEventListener('click', function() {
              showMacroCommandToggle.classList.toggle('active');
            });
          }

          const performTlsAfterHomeToggle = document.getElementById('rcs-perform-tls-after-home-toggle');
          if (performTlsAfterHomeToggle) {
            performTlsAfterHomeToggle.addEventListener('click', function() {
              performTlsAfterHomeToggle.classList.toggle('active');
            });
          }

          window.addEventListener('message', handleServerStateUpdate);

          // Populate number of tools dropdown (1-99)
          const numberOfToolsSelect = getInput('rcs-number-of-tools');
          if (numberOfToolsSelect) {
            for (let i = 1; i <= 98; i++) {
              const option = document.createElement('option');
              option.value = i;
              option.textContent = i;
              numberOfToolsSelect.appendChild(option);
            }
          }

          applyInitialSettings();
          updateRapidChangeState();

          registerButton(POCKET_PREFIX, 'rcs-pocket1-grab');
          registerButton(TOOLSETTER_PREFIX, 'rcs-toolsetter-grab');
          registerButton(PARKING_PREFIX, 'rcs-parking-grab');

          fetchInitialCoordinates();
          populateAuxOutputs();
        })();
      </script>
    `,
      { size: 'large', closable: false }
    );
  }, {
    icon: 'logo.png'
  });
}

export async function onUnload(ctx) {
  ctx.log('Manual ToolChanger unloading');

  // Reset tool settings to give control back to Settings > General
  const pluginSettings = ctx.getSettings() || {};
  const appSettings = ctx.getAppSettings() || {};

  try {
    const response = await fetch(`http://localhost:${resolveServerPort(pluginSettings, appSettings)}/api/settings`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tool: {
          source: null,
          count: 0,
          manual: false,
          tls: false
        }
      })
    });

    if (response.ok) {
      ctx.log('Tool settings reset: count=0, manual=false, tls=false, source=null');
    } else {
      ctx.log(`Failed to reset tool settings: ${response.status}`);
    }
  } catch (error) {
    ctx.log('Failed to reset tool settings on plugin unload:', error);
  }
}
