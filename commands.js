/**
 * Manual ToolChanger - Command Processor
 * Pure command processing logic for manual tool change workflow.
 * Runs on Node.js natively OR on .NET via Jint.
 * No import/require/fetch/ctx — pure input→output.
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

// === Constants ===

const PROBE_TOOL_NUMBER = 99;

// === M6 Pattern Matching (inlined from gcode-patterns.js) ===

const M6_PATTERN = /(?:^|[^A-Z])M0*6(?:\s*T0*(\d+)|(?=[^0-9T])|$)|(?:^|[^A-Z])T0*(\d+)\s*M0*6(?:[^0-9]|$)/i;

function isGcodeComment(command) {
  const trimmed = command.trim();
  const withoutLineNumber = trimmed.replace(/^N\d+\s*/i, '');
  if (withoutLineNumber.startsWith(';')) {
    return true;
  }
  if (withoutLineNumber.startsWith('(') && withoutLineNumber.endsWith(')')) {
    return true;
  }
  return false;
}

function parseM6Command(command) {
  if (!command || typeof command !== 'string') {
    return null;
  }
  if (isGcodeComment(command)) {
    return null;
  }
  const normalizedCommand = command.trim().toUpperCase();
  const match = normalizedCommand.match(M6_PATTERN);
  if (!match) {
    return null;
  }
  const toolNumberStr = match[1] || match[2];
  const toolNumber = toolNumberStr ? parseInt(toolNumberStr, 10) : null;
  return {
    toolNumber: Number.isFinite(toolNumber) ? toolNumber : null,
    matched: true
  };
}

// === Sanitization / Validation Helpers ===

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

  // Probe Tool Settings
  addProbe: raw.addProbe ?? false,
  probeLoadGcode: raw.probeLoadGcode ?? '',
  probeUnloadGcode: raw.probeUnloadGcode ?? '',

  // Tool Change Events
  preToolChangeGcode: raw.preToolChangeGcode ?? '',
  postToolChangeGcode: raw.postToolChangeGcode ?? '',
  abortEventGcode: raw.abortEventGcode ?? ''
});

// === Tool Offset Lookup (pure, from pre-fetched array) ===

function getToolOffsets(toolNumber, tools) {
  if (!toolNumber || toolNumber <= 0 || !Array.isArray(tools)) {
    return { x: 0, y: 0, z: 0 };
  }
  const tool = tools.find(t => t.toolNumber === toolNumber);
  if (tool && tool.offsets) {
    return { x: tool.offsets.x || 0, y: tool.offsets.y || 0, z: tool.offsets.tlsZ || 0 };
  }
  return { x: 0, y: 0, z: 0 };
}

// === G-code Generation Helpers ===

const formatGCode = (gcode) => {
  return gcode
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.length > 0);
};

// === Routine Generators ===

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
    G4 P0.2
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

function isManualTool(toolNumber, settings) {
  return toolNumber > settings.numberOfTools;
}

function createToolUnload(settings, currentTool, targetTool) {
  const useRCS = settings.autoSwap && !isManualTool(currentTool, settings);
  const messageCode = useRCS ? 'PLUGIN_MANUALTOOLCHANGE:UNLOAD_MESSAGE' : 'PLUGIN_MANUALTOOLCHANGE:UNLOAD_MESSAGE_MANUAL';
  const needsPause = settings.pauseBeforeUnload;
  const confirmationLines = needsPause ? `
    G4 P0
    (MSG, ${messageCode})
    M0` : '';

  if (useRCS) {
    const pauseSequence = needsPause ? `
      G53 G0 X${settings.parking.x} Y${settings.parking.y}
      ${confirmationLines}
      G53 G0 X${settings.pocket1.x} Y${settings.pocket1.y}` : `
      G53 G0 X${settings.pocket1.x} Y${settings.pocket1.y}`;

    const moveToManualAfterUnload = targetTool !== 0 ? `
      G53 G0 X${settings.parking.x} Y${settings.parking.y}` : '';

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

function createToolLoad(settings, toolNumber, hasUnload, currentTool) {
  const wasManualUnload = hasUnload && isManualTool(currentTool, settings);
  const useRCS = settings.autoSwap && !isManualTool(toolNumber, settings);
  const messageCode = useRCS
    ? (wasManualUnload ? `PLUGIN_MANUALTOOLCHANGE:LOAD_AFTER_MANUAL_MESSAGE_${toolNumber}` : `PLUGIN_MANUALTOOLCHANGE:LOAD_MESSAGE_${toolNumber}`)
    : (wasManualUnload ? `PLUGIN_MANUALTOOLCHANGE:SWAP_MESSAGE_MANUAL_${toolNumber}` : `PLUGIN_MANUALTOOLCHANGE:LOAD_MESSAGE_MANUAL_${toolNumber}`);

  const moveToManualLocation = hasUnload ? '' : `
      G53 G0 Z${settings.zSafe}
      G53 G0 X${settings.parking.x} Y${settings.parking.y}`;

  if (useRCS) {
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

function buildUnloadTool(settings, currentTool, targetTool) {
  if (currentTool === 0) {
    return '';
  }

  if (currentTool === PROBE_TOOL_NUMBER) {
    const probeUnloadGcode = settings.probeUnloadGcode?.trim() || '';
    if (probeUnloadGcode) {
      return `
        (Unload Probe Tool T${PROBE_TOOL_NUMBER})
        G53 G0 Z${settings.zSafe}
        ${probeUnloadGcode}
        M61 Q0
      `.trim();
    } else {
      return `
        (Unload Probe Tool T${PROBE_TOOL_NUMBER})
        G53 G0 Z${settings.zSafe}
        G53 G0 X${settings.parking.x} Y${settings.parking.y}
        G53 G0 Z0
        G4 P0
        (MSG, PLUGIN_MANUALTOOLCHANGE:MANUAL_UNLOAD_PROBE)
        M0
        M61 Q0
      `.trim();
    }
  }

  return `
    (Unload current tool T${currentTool})
    ${createToolUnload(settings, currentTool, targetTool)}
  `.trim();
}

function buildLoadTool(settings, toolNumber, tlsRoutine, hasUnload, currentTool) {
  if (toolNumber === 0) {
    return '';
  }

  if (toolNumber === PROBE_TOOL_NUMBER) {
    const probeLoadGcode = settings.probeLoadGcode?.trim() || '';
    if (probeLoadGcode) {
      return `
        (Load Probe Tool T${PROBE_TOOL_NUMBER})
        G53 G0 Z${settings.zSafe}
        M61 Q${PROBE_TOOL_NUMBER}
        ${probeLoadGcode}
        ${tlsRoutine}
      `.trim();
    } else {
      return `
        (Load Probe Tool T${PROBE_TOOL_NUMBER})
        G53 G0 Z${settings.zSafe}
        G53 G0 X${settings.parking.x} Y${settings.parking.y}
        G53 G0 Z0
        G4 P0
        (MSG, PLUGIN_MANUALTOOLCHANGE:MANUAL_LOAD_PROBE)
        M0
        M61 Q${PROBE_TOOL_NUMBER}
        ${tlsRoutine}
      `.trim();
    }
  }

  return `
    (Load new tool T${toolNumber})
    ${createToolLoad(settings, toolNumber, hasUnload, currentTool)}
    ${tlsRoutine}
  `.trim();
}

function buildToolChangeProgram(settings, currentTool, toolNumber, toolOffsets = { x: 0, y: 0 }) {
  const tlsRoutine = createToolLengthSetRoutine(settings, toolOffsets);
  const hasUnload = currentTool !== 0;

  const unloadSection = buildUnloadTool(settings, currentTool, toolNumber);
  const loadSection = buildLoadTool(settings, toolNumber, tlsRoutine, hasUnload, currentTool);

  const preToolChangeCmd = settings.preToolChangeGcode?.trim() || '';
  const postToolChangeCmd = settings.postToolChangeGcode?.trim() || '';

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

// === Command Handlers (synchronous, no host dependency) ===

function handleTLSCommand(commands, context, settings) {
  const tlsIndex = commands.findIndex(cmd =>
    cmd.isOriginal && cmd.command.trim().toUpperCase() === '$TLS'
  );

  if (tlsIndex === -1) {
    return;
  }

  const currentTool = context.machineState?.tool ?? 0;
  const toolOffsets = getToolOffsets(currentTool, context.tools);

  const tlsCommand = commands[tlsIndex];
  const toolLengthSetRoutine = createToolLengthSetRoutine(settings, toolOffsets);

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
      return {
        command: line,
        displayCommand: showMacroCommand ? null : tlsCommand.command.trim(),
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

  commands.splice(tlsIndex, 1, ...expandedCommands);
}

function handlePocket1Command(commands, settings) {
  const pocket1Index = commands.findIndex(cmd =>
    cmd.isOriginal && cmd.command.trim().toUpperCase() === '$POCKET1'
  );

  if (pocket1Index === -1) {
    return;
  }

  const pocket1Command = commands[pocket1Index];
  const gcode = `
    G53 G21 G90 G0 Z${settings.zSafe}
    G53 G21 G90 G0 X${settings.pocket1.x} Y${settings.pocket1.y}
  `.trim();

  const pocket1Program = formatGCode(gcode);
  const showMacroCommand = settings.showMacroCommand ?? false;

  const expandedCommands = pocket1Program.map((line, index) => {
    if (index === 0) {
      return {
        command: line,
        displayCommand: showMacroCommand ? null : pocket1Command.command.trim(),
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

  commands.splice(pocket1Index, 1, ...expandedCommands);
}

function handleHomeCommand(commands, context, settings) {
  const homeIndex = commands.findIndex(cmd =>
    cmd.isOriginal && cmd.command.trim().toUpperCase() === '$H'
  );

  if (homeIndex === -1) {
    return;
  }

  if (!settings.performTlsAfterHome) {
    return;
  }

  const currentTool = context.machineState?.tool ?? 0;
  const toolOffsets = getToolOffsets(currentTool, context.tools);

  const homeCommand = commands[homeIndex];
  const tlsRoutine = createToolLengthSetRoutine(settings, toolOffsets);

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

function handleM6Command(commands, context, settings) {
  const m6Index = commands.findIndex(cmd => {
    if (!cmd.isOriginal) return false;
    const parsed = parseM6Command(cmd.command);
    return parsed?.matched && parsed.toolNumber !== null;
  });

  if (m6Index === -1) {
    return;
  }

  const m6Command = commands[m6Index];
  const parsed = parseM6Command(m6Command.command);

  if (!parsed?.matched || parsed.toolNumber === null) {
    return;
  }

  const toolNumber = parsed.toolNumber;
  const currentTool = context.machineState?.tool ?? 0;
  const toolOffsets = getToolOffsets(toolNumber, context.tools);

  const toolChangeProgram = buildToolChangeProgram(settings, currentTool, toolNumber, toolOffsets);
  const showMacroCommand = settings.showMacroCommand ?? false;

  const expandedCommands = toolChangeProgram.map((line, index) => {
    if (index === 0) {
      return {
        command: line,
        displayCommand: showMacroCommand ? null : m6Command.command.trim(),
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

  commands.splice(m6Index, 1, ...expandedCommands);
}

// === Main Entry Point ===

function onBeforeCommand(commands, context, settings) {
  // Use core app's safe Z height setting, fallback to 0 (machine Z0)
  if (context && context.safeZHeight !== undefined) {
    settings.zSafe = context.safeZHeight;
  }

  handleHomeCommand(commands, context, settings);
  handleTLSCommand(commands, context, settings);
  handlePocket1Command(commands, settings);
  handleM6Command(commands, context, settings);
  return commands;
}

