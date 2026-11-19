# Third-party notices

Pagekit is MIT licensed. The runtime dependencies below are used under licenses that are compatible with MIT. Pagekit does not use the GPL option of any dual-licensed package.

## parse5 8.1.1

MIT. Copyright (c) 2013-2019 Ivan Nikulin (ifaaan@gmail.com, https://github.com/inikulin)

## entities 9.0.0

BSD-2-Clause. Copyright (c) Felix Böhm. Dependency of parse5.

## openxmlsdkts 1.0.13

MIT. Copyright (c) 2026 Eric White

## ltxmlts 1.0.13

MIT. Copyright (c) Eric White. Dependency of openxmlsdkts.

## jszip 3.10.2

JSZip is dual-licensed MIT OR GPL-3.0-or-later. Pagekit elects the MIT license and does not distribute JSZip under the GPL.

MIT copyright (c) 2009-2016 Stuart Knightley, David Duponchel, Franz Buchinger, António Afonso

## pako 1.0.11

MIT. Dependency of JSZip. Copyright (C) 2014-2017 by Vitaly Puzrin and Andrei Tuputcyn.

## ECMA-376 schemas (tests only)

`schemas/ecma-376` contains the Office Open XML transitional schemas published with ECMA-376, plus the W3C schema for the XML namespace. They are used only by tests and CI to validate generated `.docx` parts. They are not linked into the library.

Two `xsd:import` elements for the XML namespace (in `wml.xsd` and `shared-math.xsd`) have a local `schemaLocation="xml.xsd"` so `xmllint --nonet` can resolve `xml:space` without fetching it. The schema types are unchanged.

## Development tools

TypeScript, Vitest, ESLint, Vite, and jsdom are development dependencies. They are not shipped in the library.
