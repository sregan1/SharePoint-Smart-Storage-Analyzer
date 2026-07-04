'use strict';

const build = require('@microsoft/sp-build-web');
const gulp = require('gulp');
const path = require('path');
const fs = require('fs');
const webpack = require('webpack');

build.addSuppression(/Warning - sass/);

// The sp-property-pane dist bundle defines its AMD module as the component
// ID string ("f9e737b7-..._1.21.1"), not as "@microsoft/sp-property-pane".
// The SPFx loader is supposed to register the public alias, but on some
// workbench versions (SharePoint Online + local dev) that step is skipped,
// so requirejs falls back to relative-path.invalid and CSP blocks the load,
// producing the "[object Object]" error.
//
// Two-part fix:
//   1. BannerPlugin — injects a self-contained IIFE before the web part's
//      define() that registers the alias in requirejs at bundle-evaluation
//      time.  sp-property-pane is always preloaded by sp-webpart-base before
//      our bundle runs, so the component-ID module is already in the registry.
//   2. patchManifest() — ensures @microsoft/sp-property-pane is present in
//      temp/manifests.json for "gulp bundle --ship" (ship build).
function patchManifest() {
  const manifestPath = path.join(__dirname, 'temp', 'manifests.json');
  try {
    const raw = fs.readFileSync(manifestPath, 'utf8');
    const manifests = JSON.parse(raw);
    let changed = false;
    for (const m of manifests) {
      const sr = m && m.loaderConfig && m.loaderConfig.scriptResources;
      if (sr && !sr['@microsoft/sp-property-pane']) {
        sr['@microsoft/sp-property-pane'] = {
          type: 'component',
          id: 'f9e737b7-f0df-4597-ba8c-3060f82380db',
          version: '1.18.2',
        };
        changed = true;
      }
    }
    if (changed) {
      fs.writeFileSync(manifestPath, JSON.stringify(manifests));
      console.log('[patch-manifest] Injected @microsoft/sp-property-pane into scriptResources');
    }
  } catch (e) {
    // temp/manifests.json may not exist yet on first run — that is fine.
  }
}

// Webpack plugin that re-patches after EVERY webpack compilation
// (initial build AND each watch-mode rebuild during "gulp serve").
// Uses the 'done' hook so it fires after SPFx's own manifest-writing
// plugin has already written temp/manifests.json.
//
// Also injects a banner IIFE into the web part bundle that registers
// '@microsoft/sp-property-pane' in requirejs before the main define()
// runs.  The SPFx loader preloads sp-property-pane as the component ID
// module (e.g. "f9e737b7-..._1.21.1") but the SharePoint workbench's
// runtime sometimes omits the AMD alias for '@microsoft/sp-property-pane',
// causing requirejs to fall back to relative-path.invalid.
build.webpack.mergeConfig({
  plugins: [
    new webpack.BannerPlugin({
      // This IIFE runs synchronously before define(). By this point
      // sp-property-pane is already in requirejs.s.contexts._.defined
      // (loaded via sp-webpart-base's preloadComponents).  We find it
      // by the component-ID prefix and register the public alias.
      banner: `(function(){var d=self.define,r=self.require;if(typeof d!=='function')return;if(r&&r.defined&&r.defined('@microsoft/sp-property-pane'))return;var id='f9e737b7-f0df-4597-ba8c-3060f82380db_1.21.1';try{var reg=r.s.contexts._.defined;var pfx='f9e737b7-f0df-4597-ba8c-3060f82380db_';for(var k in reg){if(k.indexOf(pfx)===0){id=k;break;}}}catch(e){}d('@microsoft/sp-property-pane',[id],function(m){return m;});})();`,
      raw: true,
      entryOnly: true,
    }),
    {
      apply(compiler) {
        compiler.hooks.done.tap('PatchSPPropertyPane', () => {
          patchManifest();
        });
      },
    },
  ],
});

build.initialize(gulp);

// Also patch as a post-step of "gulp bundle --ship" in case mergeConfig
// plugins aren't applied in the ship path.
const originalBundle = gulp.task('bundle');
gulp.task('bundle', gulp.series(originalBundle, function patchAfterBundle(done) {
  patchManifest();
  done();
}));

gulp.task('serve', gulp.series('serve-deprecated'));
