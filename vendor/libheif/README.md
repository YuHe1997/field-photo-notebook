# libheif (vendored)

`libheif-bundle.esm.js` is the unmodified `libheif-wasm/libheif-bundle.mjs` from the npm package
[`libheif-js@1.19.8`](https://github.com/catdad-experiments/libheif-js) (Emscripten build of
[libheif](https://github.com/strukturag/libheif), WebAssembly embedded). It is loaded on demand by
`heic.js` only when the browser cannot decode a HEIC/HEIF photo itself.

License: GNU LGPL v3 (`LICENSE`, `LICENSE-libheif`). You may replace this file with any compatible
build of the same package version.
