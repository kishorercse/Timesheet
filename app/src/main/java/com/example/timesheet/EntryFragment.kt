package com.example.timesheet

import android.app.DatePickerDialog
import android.app.TimePickerDialog
import android.icu.util.Calendar
import android.os.Bundle
import android.os.Environment
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.Button
import android.widget.EditText
import android.widget.Toast
import androidx.fragment.app.Fragment
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

class EntryFragment : Fragment() {

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {

        val view = inflater.inflate(R.layout.entry_fragment, container, false)
        val pickDateButton = view.findViewById<Button>(R.id.pickDate)
        val empIdValueEditText = view.findViewById<EditText>(R.id.empIdValue)
        val pickTimeInButton = view.findViewById<Button>(R.id.timeInValue)
        val pickTimeOutButton = view.findViewById<Button>(R.id.timeOutValue)
        val submitButton = view.findViewById<Button>(R.id.submit)
        val openFileButton = view.findViewById<Button>(R.id.openFile)

        var dayOfWeek: String? = null

        pickDateButton.setOnClickListener {
            val calendar = Calendar.getInstance()

            val calendarYear = calendar.get(Calendar.YEAR)
            val month = calendar.get(Calendar.MONTH)
            val day = calendar.get(Calendar.DAY_OF_MONTH)

            val datePickerDialog = DatePickerDialog(
                requireContext(),
                { _, year, monthOfYear, dayOfMonth ->
                    pickDateButton.text = String.format("%02d-%02d-%d",dayOfMonth, (monthOfYear + 1), year)
                    val calendar = Calendar.getInstance()
                    calendar.set(year, monthOfYear, dayOfMonth)
                    dayOfWeek = when(calendar.get(Calendar.DAY_OF_WEEK)) {
                        Calendar.SUNDAY -> "SUNDAY"
                        Calendar.MONDAY -> "MONDAY"
                        Calendar.TUESDAY -> "TUESDAY"
                        Calendar.WEDNESDAY -> "WEDNESDAY"
                        Calendar.THURSDAY -> "THURSDAY"
                        Calendar.FRIDAY -> "FRIDAY"
                        Calendar.SATURDAY -> "SATURDAY"
                        else -> "EMPTY"
                    }
                }, calendarYear, month, day
            )
            datePickerDialog.show()
        }

        pickTimeInButton.setOnClickListener {
            val timePickerDialog = TimePickerDialog(requireContext(),
                { _, hourOfDay, minute ->
                    pickTimeInButton.text = String.format("%02d:%02d", hourOfDay, minute)
                }, 0, 0, true
            )
            timePickerDialog.show()
        }

        pickTimeOutButton.setOnClickListener {
            val timePickerDialog = TimePickerDialog(requireContext(),
                { view, hourOfDay, minute ->
                    pickTimeOutButton.text = String.format("%02d:%02d", hourOfDay, minute)

                }, 0, 0, true
            )
            timePickerDialog.show()
        }

        submitButton.setOnClickListener {

            if (empIdValueEditText.text.isNullOrEmpty()) {
                Toast.makeText(this.activity, "Please enter employee id", Toast.LENGTH_SHORT).show()
            } else if (empIdValueEditText.text.length < 3) {
                Toast.makeText(this.activity, "Please enter valid employee id", Toast.LENGTH_SHORT).show()
            } else {
                val dateValue = pickDateButton.text.toString()
                val timeInValue = pickTimeInButton.text.toString()
                val timeOutValue = pickTimeOutButton.text.toString()
                val logTime = calculateTimeDifference(timeInValue, timeOutValue, false)
                val regularWorkingTime =
                    if (dayOfWeek == "SUNDAY")
                        "00:00"
                    else if (logTime.substring(0,2).toInt() < 8)
                        "${logTime.substring(0,2)}:00"
                    else
                        "08:00"

                val logHours = logTime.substring(0,2).toInt()
                val otHours = if (logHours > 8) {
                    calculateTimeDifference(regularWorkingTime, logTime, true)
                } else if (logHours < 8) {
                    calculateTimeDifference( "08:00", regularWorkingTime, true)
                } else {
                    "00:00"
                }

                val path =
                    context?.getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS)?.absolutePath
                val file = File(path, "Timesheet.xlsx")

                if (!file.exists()) {
                    val workbook = XSSFWorkbook()
                    val fileOut = FileOutputStream(file)
                    workbook.write(fileOut)
                    fileOut.close()

                    workbook.close()
                }

                val workbook = XSSFWorkbook(FileInputStream(file))

                val sheetName = String.format("%03d", empIdValueEditText.text.toString().toInt())

                var sheet = workbook.getSheet(sheetName)

                if (sheet == null) {
                    sheet = workbook.createSheet(sheetName)

                    val row = sheet.createRow(0)
                    row.createCell(0).setCellValue("Date")
                    row.createCell(1).setCellValue("Day")
                    row.createCell(2).setCellValue("Time In")
                    row.createCell(3).setCellValue("Time Out")
                    row.createCell(4).setCellValue("Log hours")
                    row.createCell(5).setCellValue("Regular working hours")
                    row.createCell(6).setCellValue("Ot hours")

                    row.createCell(12).setCellValue("Regular working hours total")
                    row.createCell(14).setCellValue("Ot hours total")
                }

                var row = sheet.createRow(sheet.lastRowNum + 1)
                row.createCell(0).setCellValue(dateValue)
                row.createCell(1).setCellValue(dayOfWeek)
                row.createCell(2).setCellValue(timeInValue)
                row.createCell(3).setCellValue(timeOutValue)
                row.createCell(4).setCellValue(logTime)
                row.createCell(5).setCellValue(regularWorkingTime)
                row.createCell(6).setCellValue(otHours)


                val fileOut = FileOutputStream(file)
                workbook.write(fileOut)
                fileOut.close()

                workbook.close()

                Toast.makeText(context, "Entry added", Toast.LENGTH_SHORT).show()
                empIdValueEditText.text.clear()

            }
        }

        openFileButton.visibility = View.GONE

//        openFileButton.setOnClickListener {
//            val path =
//                context?.getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS)?.absolutePath
//            val file = File(path, "Timesheet.xlsx")
//            val intent = Intent(Intent.ACTION_VIEW)
//            intent.flags = Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION
//            val fileUri = FileProvider.getUriForFile(requireContext(), requireContext().applicationContext.packageName + ".provider", file)
//
//            intent.setDataAndType(fileUri, "application/vnd.ms-excel")
//            startActivity(intent)
//        }
        return view
    }

    private fun calculateTimeDifference(timeIn: String, timeOut:String, deductHours: Boolean): String {
        val startTimeArray = timeIn.split(":")
        val endTimeArray = timeOut.split(":")

        val startHours = startTimeArray[0].toInt()
        val startMinutes = startTimeArray[1].toInt()
        var endHours = endTimeArray[0].toInt()
        val endMinutes = endTimeArray[1].toInt()

        if (!deductHours && endHours < startHours) {
            endHours += 24
        }

        val startTime = startHours * 60 + startMinutes
        val endTime = endHours * 60 + endMinutes

        val diff = endTime - startTime

        val diffHours = diff / 60
        var diffMinutes = diff % 60
        if (deductHours) {
            diffMinutes = 0
        }

        return String.format("%02d:%02d",diffHours, diffMinutes)
    }
}