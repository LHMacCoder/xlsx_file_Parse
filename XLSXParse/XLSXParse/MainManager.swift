//
//  MainManager.swift
//  XLSXParse
//
//  Created by 浩  林 on 2019/7/7.
//  Copyright © 2019 linhao. All rights reserved.
//

import Cocoa
import CoreXLSX

class MainManager: NSObject,NSTableViewDelegate,NSTableViewDataSource{
    
    var columnTitles = [String]()
    var parseStrings = [String]()
    var duplicateStrings = [String]()
    var duplicateKeys = Set<String>()
    
    var currentSheet: String?
    
    @IBOutlet var detailText: NSTextView!
    @IBOutlet weak var filePath: NSTextField!
    @IBOutlet weak var keyValue: NSTextField!
    @IBOutlet weak var tableView: NSTableView!
    @IBOutlet weak var whetherFilter: NSButton!
    
    @IBAction func inputAction(_ sender: Any) {
        let savePanel = NSOpenPanel()
        savePanel.allowedFileTypes = ["xlsx"]
        savePanel.canCreateDirectories = false
        savePanel.runModal()
        filePath.stringValue = savePanel.url?.path ?? ""
    }
    
    @IBAction func exportAction(_ sender: Any) {
        //                // 创建保存的文件
        //                if (!mutableString.isEmpty){
        //                    var text = ""
        //                    let filePath = "\(NSHomeDirectory())/Documents/\(fileName).txt"
        //                    if (!FileManager.default.fileExists(atPath: filePath)){
        //                        FileManager.default.createFile(atPath: filePath, contents: nil, attributes: nil)
        //                    }
        //                    else{
        //                        text = try! String(contentsOf:URL(fileURLWithPath: filePath) , encoding: .utf8)
        //                    }
        //                    text += mutableString
        //
        //                    try! text.write(to: URL(fileURLWithPath: filePath), atomically: false, encoding: .utf8)
        //
        //                }
    }
    
    @IBAction func parseAction(_ sender: Any) {
        if (keyValue.stringValue.isEmpty){
            return;
        }
        guard let file = XLSXFile(filepath: filePath.stringValue) else {
            fatalError("XLSX file corrupted or does not exist")
        }
        
        for path in try! file.parseWorksheetPaths() {
            let sheetPath = NSString(string:path)
            currentSheet = sheetPath.lastPathComponent
            
            let sharedStrings = try! file.parseSharedStrings()
            let ws = try! file.parseWorksheet(at: path)
            
            // 获取所有列的标识
            var columnReferences = [String]()
            let row = ws.data?.rows[0]
            for cell in row?.cells ?? []{
                let reference = cell.reference.column.value
                columnReferences.append(reference)
            }
            
            // 首先获取作为key值的列
            var keyStrings = ws.cells(atColumns: [ColumnReference(keyValue.stringValue)!])
                .filter { $0.type == "s" }
                .compactMap { $0.value }
                .compactMap { Int($0) }
                .compactMap { sharedStrings.items[$0].text }
            
            
            // 获取重复的key
            duplicateKeys = keyStrings.getDuplicates({$0})
            
            // 过滤重复的key
            if (whetherFilter.state == .on){
                keyStrings = keyStrings.filterDuplicates({$0})
            }
            
            // 循环获取其他列
            for index in 0..<columnReferences.count{
                var mutableString = "\(currentSheet!)    "
                var duplicateString = ""
                var fileName = ""
                var temp = [String]()
                
                let ref = columnReferences[index]
                if (ref == keyValue.stringValue){
                    for index in 0..<keyStrings.count{
                        mutableString += "\"\(keyStrings[index].replacingOccurrences(of: "\"", with: "\\\""))\"" + " = " + "\"\(keyStrings[index].replacingOccurrences(of: "\"", with: "\\\""))\";" + "\n"
                    }
                    fileName = keyStrings[0]
                    addColumnTitles(fileName, atIndex: index)
                    addParseString(mutableString, atIndex: index)
                    duplicateStrings.append("")
                }
                else{
                    var columnCStrings = ws.cells(atColumns: [ColumnReference(ref)!])
                        .filter { $0.type == "s" }
                        .compactMap { $0.value }
                        .compactMap { Int($0) }
                        .compactMap { sharedStrings.items[$0].text }
                    // 过滤重复的key
                    if (whetherFilter.state == .on){
                        columnCStrings = columnCStrings.filterDuplicates({$0})
                    }
                    
                    if (columnCStrings.count == keyStrings.count){
                        var sortArray = [String]()
                        var duplicateArray = [String]()
                        
                        for index in 0..<keyStrings.count{
                            mutableString += "\"\(keyStrings[index].replacingOccurrences(of: "\"", with: "\\\""))\"" + " = " + "\"\(columnCStrings[index].replacingOccurrences(of: "\"", with: "\\\""))\";" + "\n"
                            
                            let containKey = temp.contains(keyStrings[index])
                            let containValue = temp.contains(columnCStrings[index])
                            if (containKey || containValue){
                                let str1 = "第\(index + 1)行：\"\(keyStrings[index].replacingOccurrences(of: "\"", with: "\\\""))\"" + " = " + "\"\(columnCStrings[index].replacingOccurrences(of: "\"", with: "\\\""))\";" + "\n"
                                if (!sortArray.contains(str1)){
                                    sortArray.append(str1)
                                }
                                
                                // 第一个重复元素
                                if (containKey){
                                    if let i = keyStrings.firstIndex(of: keyStrings[index]){
                                        let str2 = "第\(i + 1)行：\"\(keyStrings[i].replacingOccurrences(of: "\"", with: "\\\""))\"" + " = " + "\"\(columnCStrings[i].replacingOccurrences(of: "\"", with: "\\\""))\";" + "\n"
                                        if (!duplicateArray.contains(str2)){
                                            duplicateArray.append(str2)
                                        }
                                    }
                                }
                                else{
                                    if let i = columnCStrings.firstIndex(of: columnCStrings[index]){
                                        let str2 = "第\(i + 1)行：\"\(keyStrings[i].replacingOccurrences(of: "\"", with: "\\\""))\"" + " = " + "\"\(columnCStrings[i].replacingOccurrences(of: "\"", with: "\\\""))\";" + "\n"
                                        if (!duplicateArray.contains(str2)){
                                            duplicateArray.append(str2)
                                        }
                                    }
                                }
                                
                            }
                            else{
                                temp.append(keyStrings[index])
                                temp.append(columnCStrings[index])
                            }
                        }
                        sortArray += duplicateArray
                        sortArray = sortArray.sorted(by: { (s1, s2) -> Bool in
                            let index1 = s1.firstIndex(of: "：") ?? s1.startIndex
                            let index2 = s2.firstIndex(of: "：") ?? s2.startIndex
                            let ss1 = s1[index1...]
                            let ss2 = s2[index2...]
                            return ss1 < ss2
                        })
                        
                        duplicateString = sortArray.joined()
                        fileName = columnCStrings[0]
                        addColumnTitles(fileName, atIndex: index)
                        addParseString(mutableString, atIndex: index)
                        duplicateStrings.append(duplicateString)
                    }
                    else{
                        if (columnCStrings.count > 0){
                            addColumnTitles(columnCStrings[0], atIndex: index)
                            
                            let duplicateStrings = """
                            \(currentSheet!)
                            该项数量和key值数量不相等，检查该项数量或者以下重复的key是否有不同的翻译!
                            
                            \(duplicateKeys)
                            """
                            addParseString(duplicateStrings, atIndex: index)
                            self.duplicateStrings.append("")
                        }
                    }
                }
            }
            
        }
        tableView.reloadData()
    }
    
    
    func addColumnTitles(_ title:String, atIndex index:Int){
        if (columnTitles.count == 0){
            columnTitles.append(title)
        }
        else if (index >= columnTitles.count){
            columnTitles.append(title)
        }
    }
    
    func addParseString(_ string:String,atIndex index:Int){
        if (parseStrings.count == 0){
            parseStrings.append(string)
        }
        else if (index >= parseStrings.count){
            parseStrings.append(string)
        }
        else if (index < parseStrings.count){
            parseStrings[index] = """
            \(parseStrings[index])\n\n
            \(string)
            """
        }
    }
    
    @IBAction func clearAction(_ sender: Any) {
        columnTitles.removeAll()
        parseStrings.removeAll()
        duplicateStrings.removeAll()
        duplicateKeys.removeAll()
        tableView.reloadData()
        detailText.textStorage?.setAttributedString(NSMutableAttributedString(string: ""))
    }
    
    func numberOfRows(in tableView: NSTableView) -> Int {
        return columnTitles.count
    }
    
    func tableView(_ tableView: NSTableView, viewFor tableColumn: NSTableColumn?, row: Int) -> NSView? {
        let view : NSTableCellView = tableView.makeView(withIdentifier: tableColumn?.identifier ?? NSUserInterfaceItemIdentifier(rawValue: ""), owner: nil) as! NSTableCellView
        view.textField?.stringValue = columnTitles[row]
        return view
    }
    
    func tableViewSelectionDidChange(_ notification: Notification) {
        let selectedRow = tableView.selectedRow
        if (selectedRow == NSNotFound){
            return;
        }
        let attributedString = NSMutableAttributedString(string: parseStrings[selectedRow] + "\n\n*****************以下为重复key的翻译************************\n\n" +
            duplicateStrings[selectedRow])
        attributedString.addAttributes([NSAttributedString.Key.font : NSFont.systemFont(ofSize:14.0),NSAttributedString.Key.foregroundColor:NSColor(srgbRed: 203/255.0, green: 75/255.0, blue: 22/255.0, alpha: 1.0)], range: NSMakeRange(0, attributedString.length))
        detailText.textStorage?.setAttributedString(attributedString)
    }
    
}

extension Array {
    // 去重
    func filterDuplicates<E: Equatable>(_ filter: (Element) -> E) -> [Element] {
        var result = [Element]()
        for value in self {
            let key = filter(value)
            if !result.map({filter($0)}).contains(key) {
                result.append(value)
            }
        }
        return result
    }
    
    // 获取相同的元素
    func getDuplicates<E :Equatable>(_ filter: (Element) -> E) -> Set<String> {
        var result = [Element]()
        var duplicateKeys = Set<String>()
        for value in self {
            let key = filter(value)
            if !result.map({filter($0)}).contains(key) {
                result.append(value)
            }
            else{
                duplicateKeys.insert(key as! String)
            }
        }
        return duplicateKeys
    }
}

