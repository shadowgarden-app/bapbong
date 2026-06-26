import { Component } from '@angular/core';
import { RouterLink, RouterLinkActive, RouterOutlet } from '@angular/router';

/** Root shell: a thin nav + the routed view (editor playground at `/`, the
 *  read-only document preview at `/preview`). */
@Component({
  selector: 'app-root',
  imports: [RouterOutlet, RouterLink, RouterLinkActive],
  template: `
    <nav class="app-nav">
      <a routerLink="/" routerLinkActive="active" [routerLinkActiveOptions]="{ exact: true }">Editor</a>
      <a routerLink="/preview" routerLinkActive="active">Preview</a>
    </nav>
    <router-outlet />
  `,
  styles: [
    `
      .app-nav {
        display: flex;
        gap: 4px;
        padding: 10px 20px;
        border-bottom: 1px solid #ececef;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif;
      }
      .app-nav a {
        padding: 6px 12px;
        border-radius: 8px;
        color: #444;
        text-decoration: none;
        font-size: 14px;
      }
      .app-nav a:hover {
        background: #f4f4f6;
      }
      .app-nav a.active {
        background: #e6f1fb;
        color: #0c447c;
        font-weight: 500;
      }
    `,
  ],
})
export class App {}
