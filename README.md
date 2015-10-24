# SpreadsheetWrapper - An abstraction layer over some APIs for Excel or Calc

## Copyright
SpreadsheetWrapper - An abstraction layer over some APIs for Excel or Calc
Copyright (C) 2015  J. Férard

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.

## Goal
 The goal of SpreasheetWrapper is to provide an abstraction layer over some APIs used to write spreadsheet documents (POI, odftoolkit, ...). It's useful when you can't choose which API will be embedded in the jar.

## Installation
```
$> mkdir spreadsheetwrapper-parent
$> git clone https://github.com/jferard/spreadsheetwrapper
$> mvn clean install
```

## Usage
In your POM :
```
<project>
	...
	<dependencies>
		...
		<dependency>
			<groupId>com.github.jferard</groupId>
			<artifactId>spreadsheetwrapper</artifactId>
			<version>1.0.1-SNAPSHOT</version>
		</dependency>

		<dependency>
			<!-- one of the supported APIs -->
			<groupId>org.odftoolkit</groupId>
			<artifactId>odfdom-java</artifactId>
			<version>0.8.7</version>
		</dependency>


	</dependencies>
</project>
```

In your code :
```
final DocumentFactoryManager manager = new DocumentFactoryManager(null);
SpreadsheetDocumentFactory factory = manager.getFactory("ods.odfdom.OdsOdfdomDocumentFactory");
final SpreadsheetDocumentWriter documentWriter = factory.create();
final SpreadsheetWriter newSheet = documentWriter.addSheet("0");
newSheet.setInteger(0, 0, 1);
...
```