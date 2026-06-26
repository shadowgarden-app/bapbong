import { Routes } from '@angular/router';
import { EditorPlayground } from './editor-playground';
import { Preview } from './preview';

export const routes: Routes = [
  { path: '', component: EditorPlayground },
  { path: 'preview', component: Preview },
];
