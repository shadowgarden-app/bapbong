import { bootstrapApplication } from '@angular/platform-browser';
import { installScrollbars } from '@shadow-garden/bapbong-ui';
import { appConfig } from './app/app.config';
import { App } from './app/app';

// The shared bapbong-ui scrollbar (idle-hidden overlay rail), so the canvas,
// the dialogs' lists and the pickers' grids scroll the same way here as in
// the desktop shell.
installScrollbars();

bootstrapApplication(App, appConfig).catch((err) => console.error(err));
