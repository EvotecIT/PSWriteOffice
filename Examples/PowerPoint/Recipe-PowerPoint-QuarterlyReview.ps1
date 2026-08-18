$path = '.\Quarterly-Business-Review.pptx'
$trend = @(
    [pscustomobject]@{ Quarter = 'Q1'; Revenue = 4.2; Margin = 31 }
    [pscustomobject]@{ Quarter = 'Q2'; Revenue = 4.8; Margin = 33 }
    [pscustomobject]@{ Quarter = 'Q3'; Revenue = 5.1; Margin = 35 }
    [pscustomobject]@{ Quarter = 'Q4'; Revenue = 5.7; Margin = 36 }
)
$priorities = @(
    [pscustomobject]@{ Priority = 'Customer onboarding'; Owner = 'Product'; Target = 'Reduce setup time by 25%' }
    [pscustomobject]@{ Priority = 'Renewal risk'; Owner = 'Sales'; Target = 'Review top 20 accounts' }
    [pscustomobject]@{ Priority = 'Delivery capacity'; Owner = 'Operations'; Target = 'Add two automation lanes' }
)

PptNew -Path $path {
    PptSlideSize -Preset Screen16x9

    PptSlide {
        PptBackground -Color '#0F172A'
        PptTitle -Title 'Quarterly Business Review'
        PptTextBox -Text 'Performance, decisions, and next-quarter priorities' -X 95 -Y 185 -Width 700 -Height 70
        PptNotes -Text 'Open with the outcome: growth continued and the team needs three decisions.'
    }

    PptSlide {
        PptTitle -Title 'Performance trend'
        PptChart -Data $trend -CategoryProperty Quarter -SeriesProperty Revenue,Margin -Type ClusteredColumn -Title 'Revenue and margin' -X 60 -Y 120 -Width 700 -Height 300
        PptNotes -Text 'Explain the margin improvement before discussing revenue.'
    }

    PptSlide {
        PptTitle -Title 'Next-quarter priorities'
        PptTable -Data $priorities -X 55 -Y 130 -Width 720 -Height 220
        PptBullets -Bullets 'Approve owners', 'Confirm targets', 'Review progress monthly' -X 80 -Y 390 -Width 650 -Height 120
        PptNotes -Text 'Close by assigning each decision to a named owner.'
    }
}
