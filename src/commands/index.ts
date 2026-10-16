/** Registers every command module. */
import './file';
import './edit';
import './format';
import './cells';
import './view';
import.meta.glob(['./*.ts', '!./index.ts', '!./registry.ts', '!./ops.ts'], { eager: true });
export { def, getCommand, allCommands, commandCount, searchCommands } from './registry';
