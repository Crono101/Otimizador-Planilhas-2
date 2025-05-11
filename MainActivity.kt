// MainActivity.kt
package com.exemplo.otimizadorplanilhas

import android.Manifest
import android.content.pm.PackageManager
import android.os.Bundle
import android.widget.Button
import android.widget.Toast
import android.widget.TextView
import android.view.LayoutInflater
import android.view.ViewGroup
import android.view.View
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream

// Tela principal do app
class MainActivity : AppCompatActivity() {

    private lateinit var recyclerView: RecyclerView
    private lateinit var adapter: PlanilhaAdapter
    private val dados = mutableListOf<List<String>>()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        // Layout programático (sem XML)
        recyclerView = RecyclerView(this).apply {
            layoutManager = LinearLayoutManager(this@MainActivity)
        }
        val btn = Button(this).apply {
            text = "Importar Planilha"
            setOnClickListener {
                val caminho = "/sdcard/Download/exemplo.xlsx"
                importarPlanilha(caminho)
            }
        }

        val layout = LinearLayout(this).apply {
            orientation = LinearLayout.VERTICAL
            addView(btn)
            addView(recyclerView)
        }

        setContentView(layout)

        adapter = PlanilhaAdapter(dados)
        recyclerView.adapter = adapter

        // Solicitar permissões
        if (ActivityCompat.checkSelfPermission(
                this,
                Manifest.permission.READ_EXTERNAL_STORAGE
            ) != PackageManager.PERMISSION_GRANTED
        ) {
            ActivityCompat.requestPermissions(
                this,
                arrayOf(Manifest.permission.READ_EXTERNAL_STORAGE),
                1
            )
        }
    }

    private fun importarPlanilha(caminho: String) {
        try {
            val file = File(caminho)
            val fis = FileInputStream(file)
            val workbook = XSSFWorkbook(fis)
            val sheet = workbook.getSheetAt(0)

            dados.clear()
            for (row in sheet) {
                val linha = mutableListOf<String>()
                for (cell in row) {
                    linha.add(cell.toString())
                }
                dados.add(linha)
            }

            adapter.notifyDataSetChanged()
            workbook.close()
            fis.close()
            Toast.makeText(this, "Planilha carregada com sucesso!", Toast.LENGTH_SHORT).show()
        } catch (e: Exception) {
            Toast.makeText(this, "Erro: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
}

// Adaptador para mostrar os dados da planilha
class PlanilhaAdapter(private val dados: List<List<String>>) :
    RecyclerView.Adapter<PlanilhaAdapter.LinhaViewHolder>() {

    class LinhaViewHolder(itemView: View) : RecyclerView.ViewHolder(itemView) {
        val texto: TextView = itemView.findViewById(android.R.id.text1)
    }

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): LinhaViewHolder {
        val view = LayoutInflater.from(parent.context)
            .inflate(android.R.layout.simple_list_item_1, parent, false)
        return LinhaViewHolder(view)
    }

    override fun onBindViewHolder(holder: LinhaViewHolder, position: Int) {
        holder.texto.text = dados[position].joinToString(" | ")
    }

    override fun getItemCount(): Int = dados.size
}
